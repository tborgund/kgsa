using FileHelpers;
using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;


namespace KGSA
{
    public partial class Service
    {
        FormMain main;
        public ServiceStatus status = new ServiceStatus();
        public ServiceStat statLatest = new ServiceStat();
        private Random random = new Random();
        public DateTime dbServiceDatoFra = FormMain.rangeMin;
        public DateTime dbServiceDatoTil = FormMain.rangeMin;
        public DateTime dbServiceDato = FormMain.rangeMin;
        private float GraphMaxDay = 0;

        public Service()
        {
        }

        public void Load(FormMain form)
        {
            this.main = form;
            RunPreLoadData();
            GenerateStat();
        }

        public void GenerateServiceEssentials(List<string> doc)
        {
            try
            {
                if (statLatest != null)
                {
                    var hashId = random.Next(999, 99999);

                    doc.Add("<div class='toolbox hidePdf'>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                    doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                    doc.Add("</div>");

                    doc.Add("<div class='toggleAll' id='" + hashId + "'><span class='Subtext'>Siste 60 dagers periode (oppdatert " + dbServiceDato.ToString("dddd d. MMMM yyyy HH:mm", FormMain.norway) + ") har følgende statistikk:");
                    doc.Add("<br>Totalt " + statLatest.totalt + " mottatte servicer hvor av " + statLatest.aktive + " er uferdig. " + statLatest.ferdig + " er ferdig.");
                    doc.Add("<br>" + (statLatest.over14prosent * 100).ToString("0.00") + " % (" + statLatest.over14 + " servicer) i perioden var over 14 dager og total TAT (turn-around-time) er " + statLatest.tat60.ToString("0.00") + ".");
                    doc.Add("<br>Gjenomsnittstid fra mottat til i arbeid er " + statLatest.tilarbeid60.ToString("0.00") + " dager.</span></div>");
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void GenerateServiceOversikt(List<string> doc, BackgroundWorker bw, DateTime dateArg)
        {
            try
            {
                if (dbServiceDatoFra == dbServiceDatoTil)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen servicer funnet.</span><br>");
                    return;
                }
                var hashId = random.Next(999, 99999);

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                DataTable dt = GetServiceData(dateArg, 60);
                SaveImage(dt, FormMain.settingsPath + @"\graphServiceGraph.png", main.appConfig.graphResX, (int)(main.appConfig.graphResY * 1.5D));
                doc.Add("<div class='toggleAll' id='" + hashId + "'><img src='graphServiceGraph.png' class='image' style='max-width:100%;width:" + main.appConfig.graphWidth + "px;height:auto;'></div>");

            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void GenerateServiceHistoryGraph(List<string> doc, BackgroundWorker bw, DateTime dateArg)
        {
            try
            {
                if (dbServiceDatoFra == dbServiceDatoTil)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen servicer funnet.</span><br>");
                    return;
                }

                DataTable dt = GetServiceHistory(main.appConfig.serviceHistoryDays);
                SaveImageHistory(dt, FormMain.settingsPath + @"\graphServiceGraphHistory.png", main.appConfig.graphResX, main.appConfig.graphResY);
                doc.Add("<img src='graphServiceGraphHistory.png' class='image' style='max-width:100%;width:" + main.appConfig.graphWidth + "px;height:auto;'>");

            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void GenerateServiceFerdigStats(List<string> doc, BackgroundWorker bw, DateTime dateArg)
        {
            try
            {
                var hashId = random.Next(999, 99999);

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderFerdigStats(doc);
                doc.Add("<tbody>");

                string SQL = "SELECT tblService.Ordrenr, tblServiceLogg.Tekst, tblService.Selgerkode " +
                    "FROM tblService " +
                    "INNER JOIN tblServiceLogg " +
                    "ON tblService.ServiceID=tblServiceLogg.ServiceID " +
                    "WHERE (tblService.DatoMottat >= '" + dbServiceDatoTil.AddDays(-main.appConfig.serviceHistoryDays).ToString("yyy-MM-dd") + "') AND (tblService.DatoMottat <= '" + dbServiceDatoTil.ToString("yyy-MM-dd") + "') AND tblService.Avdeling = " + main.appConfig.Avdeling;

                DataTable sqlce = main.database.GetSqlDataTable(SQL);

                List<object[]> olst = new List<object[]>();

                for (int i = 0; i < sqlce.Rows.Count; i++)
                {
                    string skRaw = sqlce.Rows[i][1].ToString();
                    if (skRaw.Contains("-> 9"))
                    {
                        int index = skRaw.IndexOf(' ', 0);
                        string sk = skRaw.Substring(0, index);
                        olst.Add(new object[] { sk, 1, 0 });
                    }
                }

                List<object[]> newList = olst
                    .GroupBy(o => o[0].ToString())
                    .Select(i => new object[]
                    {
                        i.Key,
                        i.Sum(x => (int)x[1]),
                        0
                    })
                    .ToList();

                DataView view = new DataView(sqlce);
                DataTable distinctValues = view.ToTable(true, "Ordrenr", "Selgerkode");
                for (int i = 0; i < newList.Count; i++)
                    newList[i][2] = distinctValues.AsEnumerable().Where(x => x[1].ToString() == newList[i][0].ToString()).ToList().Count;

                List<object[]> SortedList = newList.OrderByDescending(o => (int)o[1]).ToList();

                for (int i = 0; i < SortedList.Count; i++)
                {
                    if (i > main.appConfig.serviceFerdigServiceStatsAntall - 1)
                        break;
                    doc.Add("<tr><td class='text-cat'>" + main.salesCodes.GetNavn(SortedList[i][0].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + SortedList[i][2] + "</td>");
                    doc.Add("<td class='numbers-gen'>" + SortedList[i][1] + "</td>");
                    doc.Add("</tr>");
                    
                }
                doc.Add("</tbody></table></td></tr></table>");
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void GenerateFavorittAvdelinger(List<string> doc, BackgroundWorker bw, DateTime dateArg)
        {
            try
            {
                var hashId = random.Next(999, 99999);

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderHistory(doc);
                doc.Add("<tbody>");

                for (int d = 0; d < FormMain.Favoritter.Count; d++)
                {

                    DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblServiceHistory WHERE Avdeling = '" + FormMain.Favoritter[d] + "' AND Dato = '" + dbServiceDatoTil.ToString("yyy-MM-dd") + "'");

                    if (sqlce.Rows.Count == 0)
                        continue;

                    doc.Add("<tr><td class='text-cat'>" + main.avdeling.Get(sqlce.Rows[0][1].ToString()) + "</td>"); // avdeling
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[0][3] + "</td>"); // totalt
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[0][4] + "</td>"); // aktive
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[0][5] + "</td>"); // ferdig
                    doc.Add("<td class='numbers-service'>" + TatTall(sqlce.Rows[0][6].ToString()) + "</td>"); // tat
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[0][7] + "</td>"); // over14
                    doc.Add("<td class='numbers-percent'>" + Percent(sqlce.Rows[0][8].ToString()) + "</td>"); // over14prosent
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[0][9] + "</td>"); // over21
                    doc.Add("<td class='numbers-percent'>" + Percent(sqlce.Rows[0][10].ToString()) + "</td>"); // over21prosent
                    doc.Add("<td class='numbers-gen'>" + TatTall(sqlce.Rows[0][11].ToString()) + "</td>"); // tilarbeid
                    doc.Add("</tr>");

                }
                doc.Add("</tbody></table></td></tr></table>");
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void GenerateServiceList(string statusFilter, string loggFilter, List<string> doc, BackgroundWorker bw)
        {
            try
            {
                if (dbServiceDatoFra == dbServiceDatoTil)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen servicer funnet.</span><br>");
                    return;
                }

                DataTable dt;

                if (!String.IsNullOrEmpty(loggFilter))
                    dt = GetServiceSpecialList(statusFilter, loggFilter);
                else
                    dt = GetServiceList(statusFilter);

                if (dt.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle'>Ingen servicer funnet.</span><br>");
                    return;
                }

                var hashId = random.Next(999, 99999);

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderRapport(doc);
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string str = "";
                    if ((bool)dt.Rows[i]["Egenservice"])
                        str = "egenService";

                    if ((bool)dt.Rows[i]["FerdigBehandlet"])
                        str += " behandletService";

                    doc.Add("<tr class='" + str + "'><td class='text-cat'><a href='#service" + dt.Rows[i]["ServiceID"] + "'>" + dt.Rows[i]["OrdreNr"] + "</a></td>");
                    doc.Add("<td class='text-cat'>" + Convert.ToDateTime(dt.Rows[i]["DatoMottat"]).ToShortDateString() + "</td>");
                    doc.Add("<td class='text-cat'>" + dt.Rows[i]["Status"].ToString() + "</td>");
                    doc.Add("<td class='numbers-small'>" + ColorDager((int)dt.Rows[i]["Dager"]) + "</td>");
                    doc.Add("<td class='text-cat'>" + dt.Rows[i]["Navn"].ToString() + "</td>");
                    doc.Add("<td class='text-cat'>" + dt.Rows[i]["Selgerkode"].ToString() + "</td>");
                    doc.Add("<td class='text-cat'>" + dt.Rows[i]["Verksted"].ToString() + "</td>");
                    doc.Add("<td class='numbers-small'>" + Behandlet(dt.Rows[i]["FerdigBehandlet"].ToString(), (int)dt.Rows[i]["ServiceID"]) + "</td>");
                    doc.Add("</tr>");
                }

                doc.Add("</tfoot></table></td></tr></table>");
                doc.Add("<span class='Subtitle'>Servicer: " + dt.Rows.Count + "</span><br>");

            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private string Behandlet(string arg, int id)
        {
            try
            {
                if (arg == "True")
                    return "<a class=\"behandlet\" data='" + id + "' href='#'>Ja</a>";
                else
                    return "<a class=\"behandlet\" data='" + id + "' href='#'>Nei</a>";
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return "";
            }
        }

        public void GenerateServiceDetails(List<string> doc, int serviceID)
        {
            try
            {
                if (dbServiceDatoFra == dbServiceDatoTil)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen service funnet.</span><br>");
                    return;
                }

                var hashId = random.Next(999, 99999);

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblServiceLogg WHERE (ServiceID = '" + serviceID + "') ORDER BY DatoTid ASC");

                if (sqlce.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen logg funnet for angitt service.</span><br>");
                    return;
                }

                string serviceStatus = "";
                using (DataTable sqlservice = main.database.GetSqlDataTable("SELECT * FROM tblService WHERE (ServiceID = '" + serviceID + "')"))
                {
                    if (sqlservice.Rows.Count > 0)
                    {
                        doc.Add("<span class='Subtext'>");
                        doc.Add("Service ref: <b>" + sqlservice.Rows[0]["Ordrenr"].ToString() + "</b> Navn: <b>" + sqlservice.Rows[0]["Navn"].ToString() + "</b>");
                        DateTime datoMottat = Convert.ToDateTime(sqlservice.Rows[0]["DatoMottat"]);
                        doc.Add("<br>Mottat: " + datoMottat.ToString("dddd d. MMMM", FormMain.norway));
                        DateTime datoIarbeid = Convert.ToDateTime(sqlservice.Rows[0]["DatoIarbeid"]);
                        if (datoIarbeid.Date != FormMain.rangeMin.Date)
                            doc.Add("<br>I arbeid: " + datoIarbeid.ToString("d. MMMM", FormMain.norway) + " (satt i arbeid av: " + sqlservice.Rows[0]["Selgerkode"].ToString() + ")");
                        if (!String.IsNullOrEmpty(sqlservice.Rows[0]["Verksted"].ToString()))
                            doc.Add("<br>Verksted: " + sqlservice.Rows[0]["Verksted"].ToString());
                        DateTime datoFerdig = Convert.ToDateTime(sqlservice.Rows[0]["DatoFerdig"]);
                        serviceStatus = sqlservice.Rows[0]["Status"].ToString();
                        if (datoFerdig.Date != FormMain.rangeMin.Date)
                            doc.Add("<br>Ferdig: " + datoFerdig.ToString("d. MMMM", FormMain.norway));
                        else
                            doc.Add("<br>Service er pågående. Status: " + serviceStatus);
                        DateTime datoUtlevert = Convert.ToDateTime(sqlservice.Rows[0]["DatoUtlevert"]);
                        if (datoUtlevert.Date != FormMain.rangeMin.Date)
                            doc.Add("<br>Utlevert: " + datoUtlevert.ToString("d. MMMM", FormMain.norway));

                        if (datoIarbeid.Date != FormMain.rangeMin)
                            doc.Add("<br>Tid fra mottat til arbeid: " + (datoIarbeid - datoMottat).Days + " dager");
                        if (datoIarbeid.Date != FormMain.rangeMin)
                            doc.Add("<br>Tid i arbeid: " + (dbServiceDatoTil - datoIarbeid).Days + " dager");
                        if (datoFerdig.Date != FormMain.rangeMin)
                            doc.Add("<br>Total service tid: " + (datoFerdig - datoMottat).Days + " dager");
                        else
                            doc.Add("<br>Total service tid: " + ColorDager((dbServiceDatoTil - datoMottat).Days) + " dager");
                        if (DBNull.Value != sqlservice.Rows[0]["FerdigBehandlet"])
                            doc.Add("<br>Er markert som behandlet? " + Behandlet(sqlservice.Rows[0]["FerdigBehandlet"].ToString(), serviceID));
                        else
                            doc.Add("<br>Er markert som behandlet? " + Behandlet("0", serviceID));

                        doc.Add("<br></span>");
                    }
                }

                doc.Add("<br><span class='Subtitle'>Service logg:</span>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderServiceDetails(doc);
                doc.Add("<tbody>");

                var firstdate = Convert.ToDateTime(sqlce.Rows[0][2]);
                int count = sqlce.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    var currentdate = Convert.ToDateTime(sqlce.Rows[i][2]);
                    doc.Add("<tr><td class='text-cat'>" + currentdate.ToString("d. MMMM HH:mm", FormMain.norway) + "</td>"); // dato og tid
                    doc.Add("<td class='numbers-small'>" + ColorDager((currentdate - firstdate).Days) + "</td>"); // logg kode
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][3] + "</td>"); // logg kode
                    doc.Add("<td class='text-cat'>" + sqlce.Rows[i][4] + "</td>"); // tekst
                    doc.Add("</tr>");
                    if (i == 25)
                        break;
                }
                doc.Add("</tfoot></table></td></tr></table>");
                doc.Add("<a href='" + FormMain.htmlServiceList + "'>Tilbake</a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick=\"browsepage('ServiceList', '" + serviceStatus + "')\" href=\"javascript:void(0);\">List alle med status: \"" + serviceStatus + "\"</a>&nbsp;&nbsp;&nbsp;&nbsp;");

            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        //private DataTable serviceHistory;
        public void GenerateServiceHistory(List<string> doc)
        {
            try
            {
                var hashId = random.Next(999, 99999);

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblServiceHistory WHERE (Avdeling = '" + main.appConfig.Avdeling + "') AND (Dato >= '" + dbServiceDatoTil.AddDays(-main.appConfig.serviceHistoryDays).ToString("yyy-MM-dd") + "') AND (Dato <= '" + dbServiceDatoTil.ToString("yyy-MM-dd") + "') ORDER BY Dato DESC");

                if (sqlce.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Ingen historikk lagret.</span><br>");
                    return;
                }

                //serviceHistory = sqlce;

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='OutertableNormal toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                MakeTableHeaderHistory(doc);
                doc.Add("<tbody>");

                int count = sqlce.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    doc.Add("<tr><td class='text-cat'>" + Convert.ToDateTime(sqlce.Rows[i][2]).ToString("dddd d. MMMM", FormMain.norway) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][3] + "</td>"); // totalt
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][4] + "</td>"); // aktive
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][5] + "</td>"); // ferdig
                    doc.Add("<td class='numbers-service'>" + TatTall(sqlce.Rows[i][6].ToString()) + "</td>"); // tat
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][7] + "</td>"); // over14
                    doc.Add("<td class='numbers-percent'>" + Percent(sqlce.Rows[i][8].ToString()) + "</td>"); // over14prosent
                    doc.Add("<td class='numbers-gen'>" + sqlce.Rows[i][9] + "</td>"); // over21
                    doc.Add("<td class='numbers-percent'>" + Percent(sqlce.Rows[i][10].ToString()) + "</td>"); // over21prosent
                    doc.Add("<td class='numbers-gen'>" + TatTall(sqlce.Rows[i][11].ToString()) + "</td>"); // tilarbeid
                    doc.Add("</tr>");
                    if (i == 25)
                        break;
                }
                doc.Add("</tfoot></table></td></tr></table>");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void MakeTableHeaderHistory(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Dato</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=80 >Totalt</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Aktive</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Ferdig</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=50 style='background:#80c34a;'>TAT</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Over14</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=50 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Over21</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=50 >%</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Til arbeid</td>");
            doc.Add("</tr></thead>");
        }

        private void MakeTableHeaderFerdigStats(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Selgerkode</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=120 >Satt i arbeid</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=120 >Ferdig</td>");
            doc.Add("</tr></thead>");
        }

        private void MakeTableHeaderServiceDetails(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Dato & Tid</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Dag</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Kode</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Detaljer</td>");
            doc.Add("</tr></thead>");
        }

        public void SaveImage(DataTable dtArg, string argFilename, int argX, int argY)
        {
            Bitmap graphBitmap;
            graphBitmap = DrawToBitmap(argX, argY, dtArg);
            graphBitmap.Save(argFilename, ImageFormat.Png);
            graphBitmap.Dispose();
        }

        private Bitmap DrawToBitmap(int argX, int argY, DataTable dt)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (dt.Rows.Count == 0)
                    {
                        g.DrawString("Mangler servicer!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];
                    int offsetY = 40;
                    int offsetX = 180;
                    float X = argX - offsetX;
                    float Y = argY - offsetY;
                    int pSize = dt.Rows.Count;
                    float Hstep = X / pSize;

                    float Vstep = Y / (GraphMaxDay + (GraphMaxDay / 15));

                    PointF p1, p2, p3, p4, p5, p6, p7;
                    Pen bPen = new Pen(Color.LightGray, 3);
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen1 = new Pen(Color.Black, 3);
                    Pen pen2 = new Pen(Color.Green, 6);
                    Pen pen3 = new Pen(Color.DarkGreen, 3);
                    Font fontBig = new Font("Helvetica", 20, FontStyle.Bold);
                    Font fontNormal = new Font("Helvetica", 18, FontStyle.Regular);

                    var gpBox = new GraphicsPath();
                    var gpLines14 = new GraphicsPath();
                    var gpLines21 = new GraphicsPath();
                    var gpLines30 = new GraphicsPath();
                    var gpText14 = new GraphicsPath();
                    var gpText21 = new GraphicsPath();
                    var gpText30 = new GraphicsPath();
                    var gpDivider = new GraphicsPath();
                    var gpServiceBox = new GraphicsPath();
                    var gpText = new GraphicsPath();

                    var gpTextMonth = new GraphicsPath();
                    var gpLinesMonth = new GraphicsPath();

                    float totaltAktive = 0;
                    float totaltServicer = 0;
                    float totaltAktiveOver14 = 0;
                    float totaltAktiveOver21 = 0;
                    float totaltTat = 0;
                    float totaltTilarbeid = 0;
                    float totaltFerdig = 0;
                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        int I = Convert.ToInt32(dt.Rows[d][0]);
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageService)dt.Rows[d][2];

                        var aktive = 0;

                        store.servicer = store.servicer.OrderBy(x => x.status).ToList();

                        for (int i = 0; i < store.servicer.Count; i++)
                        {
                            var sk = store.servicer[i].selgerkode;
                            var stat = store.servicer[i].status;
                            var iarbeid = store.servicer[i].Iarbeid;
                            var ferdig = store.servicer[i].Ferdig;
                            var utlevert = store.servicer[i].Utlevert;
                            var tat = store.servicer[i].tat;
                            var tilarbeid = store.servicer[i].tilarbeid;
                            var egenservice = store.servicer[i].egenservice;
                            var behandlet = store.servicer[i].ferdigbehandlet;

                            totaltServicer++;
                            if (stat < 90)
                                aktive++;
                            else
                            {
                                totaltTat += tat;
                                totaltTilarbeid += tilarbeid;
                                totaltFerdig++;
                            }

                            p5 = new PointF(X - (Hstep * I) - Hstep + 1, Y - (Vstep * i) - Vstep + 1);
                            p6 = new PointF(X - (Hstep * I) - Hstep + 3, Y - (Vstep * i) - Vstep + 3);
                            p7 = new PointF(X - (Hstep * I) - 2, Y - (Vstep * i) - 2);
                            var rect = new RectangleF(p5, new SizeF(Hstep - 2, Vstep - 2));
                            var rectAct = new RectangleF(p6, new SizeF(Hstep - 6, Vstep - 6));

                            if (!behandlet)
                            {
                                if (egenservice)
                                    g.FillRectangle(new HatchBrush(HatchStyle.ForwardDiagonal, Color.White, status.GetStatusColor(stat)), rect);
                                else
                                    g.FillRectangle(new SolidBrush(status.GetStatusColor(stat)), rect);
                            }
                            else
                                g.FillRectangle(new SolidBrush(Tint(status.GetStatusColor(stat), Color.White, 0.5M)), rect);

                            if (iarbeid == dbServiceDatoTil || ferdig == dbServiceDatoTil || utlevert == dbServiceDatoTil)
                                g.DrawLine(new Pen(Tint(status.GetStatusColor(stat), Color.Black, 0.2M), 3), p5, p7);

                            if (!behandlet)
                                g.DrawRectangle(new Pen(Tint(status.GetStatusColor(stat), Color.Black, 0.2M), 3), new Rectangle(new Point((int)(X - (Hstep * I) - Hstep + 1), (int)(Y - (Vstep * i) - Vstep + 1)), new Size((int)(Hstep), (int)(Vstep))));
                            else
                                g.DrawRectangle(new Pen(Tint(status.GetStatusColor(stat), Color.White, 0.5M), 3), new Rectangle(new Point((int)(X - (Hstep * I) - Hstep + 1), (int)(Y - (Vstep * i) - Vstep + 1)), new Size((int)(Hstep), (int)(Vstep))));

                        }

                        int index = 0;
                        index = I + (dbServiceDatoTil - firstdate).Days;
                        if (pSize <= 31 && index < 100)
                            gpText.AddString(index.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 28, new PointF((X - (Hstep * I) - Hstep + (Hstep / 5)), Y), StringFormat.GenericDefault);
                        p1 = new PointF(X - (Hstep * I) - Hstep, Y);
                        p4 = new PointF(X - (Hstep * I), Y);
                        p2 = new PointF(X - (Hstep * I) - Hstep, Y - (Vstep * aktive));
                        p3 = new PointF(X - (Hstep * I), Y - (Vstep * aktive));
                        gpBox.StartFigure();
                        gpBox.AddLine(p1, p2);
                        gpBox.AddLine(p2, p3);
                        gpBox.AddLine(p3, p4);

                        if (dbServiceDatoTil == firstdate)
                        {
                            if ((dbServiceDatoTil - date).Days > 14)
                                totaltAktiveOver14 += aktive;
                            if ((dbServiceDatoTil - date).Days > 21)
                                totaltAktiveOver21 += aktive;
                        }
                        if ((dbServiceDatoTil - date).Days == 14)
                        {
                            if (dbServiceDatoTil == firstdate)
                                gpText14.AddString(totaltAktive + " aktive", new FontFamily("Verdana"), (int)FontStyle.Regular, 28, new PointF(X - (Hstep * I) - Hstep, 0), StringFormat.GenericDefault);
                            gpLines14.AddLine(new PointF(X - (Hstep * I) - Hstep, Y + offsetY), new PointF(X - (Hstep * I) - Hstep, 0));
                        }
                        if ((dbServiceDatoTil - date).Days == 21)
                        {
                            if (dbServiceDatoTil == firstdate)
                                gpText21.AddString(totaltAktiveOver14 + " aktive", new FontFamily("Verdana"), (int)FontStyle.Regular, 28, new PointF(X - (Hstep * I) - Hstep, 0), StringFormat.GenericDefault);
                            gpLines21.AddLine(new PointF(X - (Hstep * I) - Hstep, Y + offsetY), new PointF(X - (Hstep * I) - Hstep, 0));
                        }
                        if ((dbServiceDatoTil - date).Days == 30)
                        {
                            if (dbServiceDatoTil == firstdate)
                                gpText30.AddString(totaltAktiveOver21 + " aktive", new FontFamily("Verdana"), (int)FontStyle.Regular, 28, new PointF(X - (Hstep * I) - Hstep, 0), StringFormat.GenericDefault);
                            gpLines30.AddLine(new PointF(X - (Hstep * I) - Hstep, Y + offsetY), new PointF(X - (Hstep * I) - Hstep, 0));
                        }
                        if (date.Day == 1)
                        {
                            gpTextMonth.AddString(date.ToString("MMM"), new FontFamily("Verdana"), (int)FontStyle.Regular, 32, new PointF(X - (Hstep * I), Y), StringFormat.GenericDefault);
                            gpLinesMonth.StartFigure();
                            gpLinesMonth.AddLine(new PointF(X - (Hstep * I), 0), new PointF(X - (Hstep * I), argY));
                        }

                        totaltAktive += aktive;
                    }
                    decimal tatAvg = 0;
                    if (totaltFerdig != 0)
                        tatAvg = Math.Round(Convert.ToDecimal(totaltTat / totaltFerdig), 2);
                    decimal tilarbeidAvg = 0;
                    if (totaltFerdig != 0)
                        tilarbeidAvg = Math.Round(Convert.ToDecimal(totaltTilarbeid / totaltFerdig), 2);

                    g.DrawPath(pen1, gpBox);
                    g.DrawPath(new Pen(Color.Gray, 4), gpDivider);
                    g.FillPath(new SolidBrush(Color.Black), gpTextMonth);
                    g.DrawPath(new Pen(Color.Black, 3), gpLinesMonth);

                    var gpTat = new GraphicsPath();
                    float tat_ServiceFerdige = 0, tat_ServiceTotAktive = 0, tat_Totalt = 0, tat_Now = 0, tat_Prev = 0;

                    for (int d = dt.Rows.Count - 1; d >= 0; d--)
                    {
                        int I = Convert.ToInt32(dt.Rows[d][0]);
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageService)dt.Rows[d][2];
                        store.servicer = store.servicer.OrderBy(x => x.status).ToList();
                        for (int i = 0; i < store.servicer.Count; i++)
                        {
                            var sk = store.servicer[i].selgerkode;
                            var stat = store.servicer[i].status;
                            var iarbeid = store.servicer[i].Iarbeid;
                            var ferdig = store.servicer[i].Ferdig;
                            var utlevert = store.servicer[i].Utlevert;
                            var tat = store.servicer[i].tat;

                            if (stat < 90)
                                tat_ServiceTotAktive++;
                            else
                            {
                                tat_Totalt += tat;
                                tat_ServiceFerdige++;
                            }
                        }

                        tat_Now = tat_Totalt / tat_ServiceFerdige;

                        if (main.appConfig.serviceShowTrend)
                        {
                            p6 = new PointF(X - (Hstep * I), Y - (Vstep * tat_Now));
                            p7 = new PointF(X - (Hstep * I) - Hstep, Y - (Vstep * tat_Prev));

                            Color grade = Color.Green;
                            if (tat_Now >= 10)
                                grade = Color.Orange;
                            if (tat_Now >= 14)
                                grade = Color.Red;

                            if (tat_Prev > 0 && (dbServiceDatoTil - date).Days > 30)
                                g.DrawLine(new Pen(grade, 8), p6, p7);
                        }
                        tat_Prev = tat_Now;
                    }

                    gpText.AddString(firstdate.ToString("dddd", FormMain.norway), new FontFamily("Verdana"), (int)FontStyle.Regular, 32, new PointF(X + 4, 65), StringFormat.GenericDefault);
                    gpText.AddString(firstdate.ToString("d. MMM", FormMain.norway), new FontFamily("Verdana"), (int)FontStyle.Regular, 32, new PointF(X + 4, 100), StringFormat.GenericDefault);
                    if (firstdate.Date == dbServiceDato.Date)
                        gpText.AddString(dbServiceDato.ToString("HH:mm", FormMain.norway), new FontFamily("Verdana"), (int)FontStyle.Regular, 32, new PointF(X + 4, 135), StringFormat.GenericDefault);

                    int addH = 94;
                    int addE = 37;
                    gpText.AddString("Servicer:", new FontFamily("Verdana"), (int)FontStyle.Regular, 26, new PointF(X + 10, 200 + addH), StringFormat.GenericDefault);
                    gpText.AddString(totaltServicer.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 30, new PointF(X + 25, 200 + addH + addE), StringFormat.GenericDefault);
                    addH += 94;
                    gpText.AddString("Aktive:", new FontFamily("Verdana"), (int)FontStyle.Regular, 26, new PointF(X + 10, 200 + addH), StringFormat.GenericDefault);
                    gpText.AddString(totaltAktive.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 30, new PointF(X + 25, 200 + addH + addE), StringFormat.GenericDefault);
                    addH += 94;
                    gpText.AddString("TAT:", new FontFamily("Verdana"), (int)FontStyle.Regular, 26, new PointF(X + 10, 200 + addH), StringFormat.GenericDefault);
                    gpText.AddString(tatAvg.ToString("0.00"), new FontFamily("Verdana"), (int)FontStyle.Regular, 30, new PointF(X + 25, 200 + addH + addE), StringFormat.GenericDefault);
                    addH += 94;
                    if (dbServiceDatoTil == firstdate)
                    {
                        var percent14 = CalcPercent(totaltAktiveOver14, totaltAktive) * 100;
                        var percent21 = CalcPercent(totaltAktiveOver21, totaltAktive) * 100;

                        gpText.AddString("Over 14:", new FontFamily("Verdana"), (int)FontStyle.Regular, 26, new PointF(X + 10, 200 + addH), StringFormat.GenericDefault);
                        gpText.AddString(totaltAktiveOver14.ToString() + " (" + percent14.ToString("0.0") + "%)", new FontFamily("Verdana"), (int)FontStyle.Regular, 30, new PointF(X + 25, 200 + addH + addE), StringFormat.GenericDefault);
                        addH += 94;
                        gpText.AddString("Over 21:", new FontFamily("Verdana"), (int)FontStyle.Regular, 26, new PointF(X + 10, 200 + addH), StringFormat.GenericDefault);
                        gpText.AddString(totaltAktiveOver21.ToString() + " (" + percent21.ToString("0.0") + "%)", new FontFamily("Verdana"), (int)FontStyle.Regular, 30, new PointF(X + 25, 200 + addH + addE), StringFormat.GenericDefault);
                    }

                    g.FillPath(new SolidBrush(Color.Black), gpText);
                    g.DrawPath(new Pen(Color.Green, 6), gpLines14);
                    g.FillPath(new SolidBrush(Color.Green), gpText14);
                    g.DrawPath(new Pen(Color.Orange, 6), gpLines21);
                    g.FillPath(new SolidBrush(Color.Orange), gpText21);
                    g.DrawPath(new Pen(Color.Red, 6), gpLines30);
                    g.FillPath(new SolidBrush(Color.Red), gpText30);

                    g.DrawLine(pen1, new Point(0, (int)Y), new Point(argX, (int)Y)); // understrek
                    g.DrawLine(pen1, new Point((int)X, (int)Y), new Point((int)X, 0)); // slutt strek
                    g.DrawRectangle(pen1, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1))); // ramme

                    int height = 0;
                    if (argY > 200 && pSize <= 61)
                        height += 29 * 7;

                    using (var gpSkilt = new GraphicsPath()) // skilt skygge
                    {
                        gpSkilt.AddRectangle(new Rectangle(12, 12, 307, 47 + height));
                        using (PathGradientBrush _Brush = new PathGradientBrush(gpSkilt))
                        {
                            _Brush.WrapMode = WrapMode.Clamp;
                            ColorBlend _ColorBlend = new ColorBlend(3);
                            _ColorBlend.Colors = new Color[]{Color.Transparent, 
                           Color.FromArgb(180, Color.DimGray), 
                           Color.FromArgb(180, Color.DimGray)};
                            _ColorBlend.Positions = new float[] { 0f, .1f, 1f };
                            _Brush.InterpolationColors = _ColorBlend;
                            g.FillPath(_Brush, gpSkilt);
                        }
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(8, 8, 300, 40 + height));
                    g.DrawRectangle(pen1, new Rectangle(new Point(8, 8), new Size(300, 40 + height)));

                    g.DrawString("Aktiv service", fontNormal, new SolidBrush(Color.Black), 50, 15);
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(18, 18, 20, 22));
                    g.DrawRectangle(pen1, new Rectangle(18, 18, 20, 22));

                    int add = 29;
                    if (argY > 200 && pSize <= 61)
                    {
                        g.DrawString("Venter service", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new SolidBrush(status.GetStatusColor(1)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(1), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                        add += 29;
                        g.DrawString("I arbeid internt", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new SolidBrush(status.GetStatusColor(2)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(2), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                        add += 29;
                        g.DrawString("I arbeid eksternt", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new SolidBrush(status.GetStatusColor(10)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(10), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                        add += 29;
                        g.DrawString("Ferdig", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new SolidBrush(status.GetStatusColor(90)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(90), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                        add += 29;
                        g.DrawString("Utlevert", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new SolidBrush(status.GetStatusColor(99)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(99), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                        add += 29;
                        g.DrawString("Endret status i dag", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.DrawLine(new Pen(Tint(status.GetStatusColor(99), Color.Black, 0.2M), 3), new Point(18, 18 + add), new Point(38, 35 + add));
                        add += 29;
                        g.DrawString("Egen service", fontNormal, new SolidBrush(Color.Black), 50, 15 + add);
                        g.FillRectangle(new HatchBrush(HatchStyle.ForwardDiagonal, Color.White, status.GetStatusColor(2)), new Rectangle(18, 18 + add, 20, 22));
                        g.DrawRectangle(new Pen(Tint(status.GetStatusColor(2), Color.Black, 0.2M), 3), new Rectangle(18, 18 + add, 20, 22));
                    }
                }
                catch
                {
                }
            }
            return b;
        }

        public void SaveImageHistory(DataTable dtArg, string argFilename, int argX, int argY)
        {
            Bitmap graphBitmap;
            graphBitmap = DrawToBitmapHistory(argX, argY, dtArg);
            graphBitmap.Save(argFilename, ImageFormat.Png);
            graphBitmap.Dispose();
        }

        private Bitmap DrawToBitmapHistory(int argX, int argY, DataTable dt)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;
                    int fontHeight = Convert.ToInt32(29 * dpi);
                    int boxLength = Convert.ToInt32(22 * dpi);
                    int fontSepHeight = Convert.ToInt32((fontHeight / 5) * dpi);

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (dt.Rows.Count == 0)
                    {
                        g.DrawString("Mangler servicer!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];
                    int offsetY = 40;
                    int offsetX = 100;
                    float X = argX - offsetX;
                    float Y = argY - offsetY;
                    int pSize = dt.Rows.Count;
                    float Hstep = X / pSize;
                    float HstepSub = Hstep / FormMain.Favoritter.Count;

                    float VstepA = Y / 20f;
                    float VstepB = Y / (maxGraphHistoryValue + (maxGraphHistoryValue / 10));
                    if (float.IsInfinity(VstepB))
                        VstepB = 20;

                    PointF p1, p2;
                    Pen bPen = new Pen(Color.LightGray, 3);
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen1 = new Pen(Color.Black, 3);
                    Font fontNormal = new Font("Helvetica", 18, FontStyle.Regular);

                    g.DrawString(main.avdeling.Get(main.appConfig.Avdeling) + " Service Historikk",
                        new Font("Verdana", 36, FontStyle.Bold), new SolidBrush(Color.LightGray), 400, 0);

                    var step = Y / 4;
                    g.DrawLine(bPen, new Point(0, (int)step * 3), new Point((int)X, (int)step * 3));
                    g.DrawLine(bPen, new Point(0, (int)step * 2), new Point((int)X, (int)step * 2));
                    g.DrawLine(bPen, new Point(0, (int)step), new Point((int)X, (int)step));

                    g.DrawString("5", fontNormal, bBrush, X + 6, (step * 3) - 14);
                    g.DrawString("10", fontNormal, bBrush, X + 6, (step * 2) - 14);
                    g.DrawString("15", fontNormal, bBrush, X + 6, step - 14);

                    var gpTextMonth = new GraphicsPath();
                    var gpLinesMonth = new GraphicsPath();

                    for (int j = 2; j < dt.Columns.Count; j++)
                    {
                        var gpBox = new GraphicsPath();
                        var gpTat = new GraphicsPath();
                        List<PointF> points = new List<PointF> { };
                        for (int d = 0; d < dt.Rows.Count; d++)
                        {

                            int I = Convert.ToInt32(dt.Rows[d][0]);
                            DateTime date = Convert.ToDateTime(dt.Rows[d][1]);

                            if (j == 2)
                            {
                                if (date.DayOfWeek == DayOfWeek.Sunday && Hstep > 20)
                                {
                                    p1 = new PointF(X - (Hstep * I), Y - 90);
                                    p2 = new PointF(X - (Hstep * I), Y - 60);
                                    g.DrawString(date.ToString("MMM"), fontNormal, new SolidBrush(Color.Black), p1);
                                    g.DrawString(date.ToString("dd."), fontNormal, new SolidBrush(Color.Black), p2);
                                }

                                if (date.Date == dbServiceDatoFra.Date)
                                    g.DrawLine(pen1, new PointF(X - (Hstep * I), Y), new PointF(X - (Hstep * I), 0));

                                if (date.Day == 1)
                                {
                                    gpTextMonth.AddString(date.ToString("MMM"), new FontFamily("Verdana"), (int)FontStyle.Regular, 32, new PointF(X - (Hstep * I), Y), StringFormat.GenericDefault);
                                    gpLinesMonth.StartFigure();
                                    gpLinesMonth.AddLine(new PointF(X - (Hstep * I), 0), new PointF(X - (Hstep * I), argY));
                                }

                            }
                                

                            if (!DBNull.Value.Equals(dt.Rows[d][j]))
                            {
                                

                                var store = (ServiceHistory)dt.Rows[d][j];
                                p2 = new PointF(X - (Hstep * I), Y - (VstepA * (float)store.tat));

                                gpBox.StartFigure();
                                if (HstepSub > 10)
                                    gpBox.AddRectangle(new RectangleF(new PointF(X - (Hstep * I) - (HstepSub * (j - 2) + HstepSub), Y - (VstepB * (float)store.aktive)), new SizeF(HstepSub, VstepB * store.aktive)));
                                else if (j == 2)
                                    gpBox.AddRectangle(new RectangleF(new PointF(X - (Hstep * I) - Hstep, Y - (VstepB * (float)store.aktive)), new SizeF(Hstep, VstepB * store.aktive)));
                            }

                        }

                        g.FillPath(new SolidBrush(Tint(FormMain.favColors[j - 2], Color.White, 0.5M)), gpBox);
                        g.DrawPath(pen1, gpBox);

                    }

                    for (int j = dt.Columns.Count - 1; j > 1; j--)
                    {
                        var gpBox = new GraphicsPath();
                        var gpTat = new GraphicsPath();
                        List<PointF> points = new List<PointF> { };
                        for (int d = 0; d < dt.Rows.Count; d++)
                        {

                            int I = Convert.ToInt32(dt.Rows[d][0]);
                            DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                            if (!DBNull.Value.Equals(dt.Rows[d][j]))
                            {
                                var store = (ServiceHistory)dt.Rows[d][j];
                                p2 = new PointF(X - (Hstep * I), Y - (VstepA * (float)store.tat));
                                points.Add(p2);
                            }

                        }
                        if (favTat.Length == FormMain.Favoritter.Count)
                            g.DrawLine(new Pen(FormMain.favColors[j - 2], 8), X, Y - (VstepA * (float)favTat[j - 2]), argX, Y - (VstepA * (float)favTat[j - 2]));
                        PaintLabelTall(g, favTat[j - 2], X - 10, Y, offsetX, offsetY, VstepA);


                        if (points.Count > 1)
                        {
                            gpTat.AddLines(points.ToArray());
                            g.DrawPath(new Pen(Color.White, 12), gpTat);
                            g.DrawPath(new Pen(FormMain.favColors[j - 2], 8), gpTat);
                        }
                    }

                    g.FillPath(new SolidBrush(Color.Black), gpTextMonth);
                    g.DrawPath(new Pen(Color.Black, 3), gpLinesMonth);

                    g.DrawLine(pen1, new Point(0, (int)Y), new Point(argX, (int)Y)); // understrek
                    g.DrawLine(pen1, new Point((int)X, (int)Y), new Point((int)X, 0)); // slutt strek
                    g.DrawRectangle(pen1, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1))); // ramme



                    int height = FormMain.Favoritter.Count * fontHeight + fontHeight;

                    using (var gpSkilt = new GraphicsPath()) // skilt skygge
                    {
                        gpSkilt.AddRectangle(new Rectangle(12, 12, 347, 20 + height));
                        using (PathGradientBrush _Brush = new PathGradientBrush(gpSkilt))
                        {
                            _Brush.WrapMode = WrapMode.Clamp;
                            ColorBlend _ColorBlend = new ColorBlend(3);
                            _ColorBlend.Colors = new Color[]{Color.Transparent, 
                           Color.FromArgb(180, Color.DimGray), 
                           Color.FromArgb(180, Color.DimGray)};
                            _ColorBlend.Positions = new float[] { 0f, .1f, 1f };
                            _Brush.InterpolationColors = _ColorBlend;
                            g.FillPath(_Brush, gpSkilt);
                        }
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(8, 8, 340, 15 + height)); // skilt
                    g.DrawRectangle(pen1, new Rectangle(new Point(8, 8), new Size(340, 15 + height))); // skilt ramme

                    int add = 0;
                    if (argY > 400)
                    {
                        for (int i = 0; i < FormMain.Favoritter.Count; i++)
                        {
                            if (favTat != null)
                                g.DrawString("TAT " + main.avdeling.Get(FormMain.Favoritter[i]) + " " + favTat[i].ToString("0.00"), fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            else
                                g.DrawString("TAT " + main.avdeling.Get(FormMain.Favoritter[i]), fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            g.DrawLine(new Pen(FormMain.favColors[i], 8 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                            add += fontHeight;
                        }

                        g.DrawString("Aktive servicer", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.FillRectangle(new SolidBrush(Color.White), new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                        g.DrawRectangle(pen1, new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                    }
                }
                catch
                {
                }
            }
            return b;
        }

        private void PaintLabelTall(Graphics g, decimal argTall, float X, float Y, float offsetX, float offsetY, float Vstep, float dpi = 1)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddString(argTall.ToString("0.00"), new FontFamily("Helvetica"),
            (int)FontStyle.Bold, 28f * dpi, new PointF(X + (offsetX / 4), Y - (Vstep * (float)argTall) - 30), StringFormat.GenericDefault);

            for (int i = 1; i < 8; ++i)
            {
                Pen pen = new Pen(Color.FromArgb(32, 255, 255, 255), i);
                pen.LineJoin = LineJoin.Round;
                g.DrawPath(pen, path);
                pen.Dispose();
            }

            SolidBrush brush = new SolidBrush(Color.FromArgb(0, 0, 0));
            g.FillPath(brush, path);
        }

        private decimal[] favTat;

        public void GenerateStat()
        {
            try
            {
                if (dbServiceDatoFra.Date != dbServiceDatoTil.Date && (dbServiceDatoTil - dbServiceDatoFra).Days >= 60)
                {
                    favTat = new decimal[FormMain.Favoritter.Count];
                    for (int d = 0; d < FormMain.Favoritter.Count; d++)
                    {
                        DataTable dt = GetServiceData(dbServiceDatoTil, 60, FormMain.Favoritter[d]);
                        var obj = new ServiceStat();
                        MakeStat(dt, obj, FormMain.Favoritter[d]);
                        if (d == 0)
                            this.statLatest = obj;
                        SaveStat(obj, FormMain.Favoritter[d]);
                        dt.Dispose();
                        favTat[d] = obj.tat60;
                    }
                }
                else
                    Log.d("Service: Databasen oppfyller ikke krav for lagring av historikk.");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void MakeStat(DataTable dt, ServiceStat obj, string avd)
        {
            try
            {
                if (dt.Rows.Count == 0)
                    return;

                DateTime firstdate = DateTime.Now;
                DateTime lastdate = DateTime.Now;
                firstdate = (DateTime)dt.Rows[0][1];
                lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];

                int pSize = dt.Rows.Count;

                float totaltAktive = 0;
                float totaltServicer = 0;
                float totaltAktiveOver14 = 0;
                float totaltAktiveOver21 = 0;
                float totaltTat = 0;
                float totaltTilarbeid = 0;
                float totaltFerdig = 0;
                float totalVenter = 0;
                float totalOver14 = 0;
                float totalOver21 = 0;
                float totalOver30 = 0;

                for (int d = 0; d < dt.Rows.Count; d++)
                {
                    int I = Convert.ToInt32(dt.Rows[d][0]);
                    DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                    var store = (StorageService)dt.Rows[d][2];

                    var aktive = 0;

                    store.servicer = store.servicer.OrderBy(x => x.status).ToList();

                    for (int i = 0; i < store.servicer.Count; i++)
                    {
                        var sk = store.servicer[i].selgerkode;
                        var stat = store.servicer[i].status;
                        var iarbeid = store.servicer[i].Iarbeid;
                        var ferdig = store.servicer[i].Ferdig;
                        var utlevert = store.servicer[i].Utlevert;
                        var tat = store.servicer[i].tat;
                        var tilarbeid = store.servicer[i].tilarbeid;

                        totaltServicer++;
                        if (stat < 90)
                            aktive++;
                        else
                        {
                            totaltTat += tat;
                            totaltTilarbeid += tilarbeid;
                            totaltFerdig++;
                            if (tat > 14)
                                totalOver14++;
                            if (tat > 21)
                                totalOver21++;
                            if (tat > 30)
                                totalOver30++;
                        }

                        if (stat == 1)
                            totalVenter++;
                    }

                    if (dbServiceDatoTil == firstdate)
                    {
                        if ((dbServiceDatoTil - date).Days > 14)
                            totaltAktiveOver14 += aktive;
                        if ((dbServiceDatoTil - date).Days > 21)
                            totaltAktiveOver21 += aktive;
                    }

                    totaltAktive += aktive;
                }
                decimal tatAvg = 0;
                decimal tilarbeidAvg = 0;
                float over14Avg = 0;
                float over21Avg = 0;
                float over30Avg = 0;
                if (totaltFerdig != 0)
                {
                    tatAvg = Math.Round(Convert.ToDecimal(totaltTat / totaltFerdig), 2);
                    tilarbeidAvg = Math.Round(Convert.ToDecimal(totaltTilarbeid / totaltFerdig), 2);
                    over14Avg = totalOver14 / totaltFerdig;
                    over21Avg = totalOver21 / totaltFerdig;
                    over30Avg = totalOver30 / totaltFerdig;
                }
                

                float tat_ServiceFerdige = 0, tat_ServiceTotAktive = 0, tat_Totalt = 0, tat_Now = 0, tat_Prev = 0;

                for (int d = dt.Rows.Count - 1; d >= 0; d--)
                {
                    int I = Convert.ToInt32(dt.Rows[d][0]);
                    DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                    var store = (StorageService)dt.Rows[d][2];
                    store.servicer = store.servicer.OrderBy(x => x.status).ToList();
                    for (int i = 0; i < store.servicer.Count; i++)
                    {
                        var sk = store.servicer[i].selgerkode;
                        var stat = store.servicer[i].status;
                        var iarbeid = store.servicer[i].Iarbeid;
                        var ferdig = store.servicer[i].Ferdig;
                        var utlevert = store.servicer[i].Utlevert;
                        var tat = store.servicer[i].tat;

                        if (stat < 90)
                            tat_ServiceTotAktive++;
                        else
                        {
                            tat_Totalt += tat;
                            tat_ServiceFerdige++;
                        }
                    }

                    tat_Now = tat_Totalt / tat_ServiceFerdige;
                    tat_Prev = tat_Now;
                }

                if (dbServiceDatoTil == firstdate)
                {
                    obj.avdeling = Convert.ToInt32(avd);
                    obj.tat60 = tatAvg;
                    obj.aktive = (int)totaltAktive;
                    obj.ferdig = (int)totaltFerdig;
                    obj.over14aktiv = (int)totaltAktiveOver14;
                    obj.over21aktiv = (int)totaltAktiveOver21;
                    obj.totalt = (int)totaltServicer;
                    obj.tilarbeid60 = tilarbeidAvg;
                    obj.oppdatert = dbServiceDatoTil;
                    obj.venter = (int)totalVenter;
                    obj.over14 = (int)totalOver14;
                    obj.over14prosent = over14Avg;
                    obj.over21 = (int)totalOver21;
                    obj.over21prosent = over21Avg;
                    obj.over30 = (int)totalOver30;
                    obj.over30prosent = over30Avg;

                    var percent14 = CalcPercent(totaltAktiveOver14, totaltAktive) * 100;
                    var percent21 = CalcPercent(totaltAktiveOver21, totaltAktive) * 100;
                    obj.over14aktivprosent = percent14;
                    obj.over21aktivprosent = percent21;
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private bool SaveStat(ServiceStat obj, string avd)
        {
            try
            {
                if (obj != null)
                {
                    if (obj.avdeling == 0 || obj.totalt == 0)
                        return false;

                    var command = new SqlCeCommand("DELETE FROM tblServiceHistory WHERE (Dato = '" + obj.oppdatert.ToString("yyy-MM-dd") + "') AND (Avdeling = '" + avd + "')", main.connection);
                    var deleted = command.ExecuteNonQuery();
                    Log.d("Oppdatert historikk for " + obj.avdeling + " - " + obj.oppdatert);

                    string sql = "INSERT INTO tblServiceHistory (Avdeling, Dato, Totalt, Aktive, Ferdig, Tat, Over14, Over14prosent, Over21, Over21prosent, Tilarbeid) " +
        "VALUES (@Avdeling, @Dato, @Totalt, @Aktive, @Ferdig, @Tat, @Over14, @Over14prosent, @Over21, @Over21prosent, @Tilarbeid)";

                    using (SqlCeCommand cmd = new SqlCeCommand(sql, main.connection))
                    {
                        cmd.Parameters.AddWithValue("@Avdeling", obj.avdeling);
                        cmd.Parameters.Add("@Dato", SqlDbType.DateTime).Value = obj.oppdatert;
                        cmd.Parameters.AddWithValue("@Totalt", obj.totalt);
                        cmd.Parameters.AddWithValue("@Aktive", obj.aktive);
                        cmd.Parameters.AddWithValue("@Ferdig", obj.ferdig);
                        cmd.Parameters.AddWithValue("@Tat", obj.tat60);
                        cmd.Parameters.AddWithValue("@Over14", obj.over14);
                        cmd.Parameters.AddWithValue("@Over14prosent", obj.over14prosent * 100);
                        cmd.Parameters.AddWithValue("@Over21", obj.over21);
                        cmd.Parameters.AddWithValue("@Over21prosent", obj.over21prosent * 100);
                        cmd.Parameters.AddWithValue("@Tilarbeid", obj.tilarbeid60);

                        cmd.CommandType = System.Data.CommandType.Text;
                        int result = cmd.ExecuteNonQuery();

                        if (result > 0)
                            return true;
                        else
                            return false;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return false;
            }
        }

        public DataTable GetServiceData(DateTime dateArg, int days, string avd = "")
        {
            try
            {
                DataTable dt = ReadyTableGraphAdvanced();
                DateTime DatoFra = dateArg.AddDays(-days);
                DateTime checkValue = new DateTime(2001, 1, 1);

                if (String.IsNullOrEmpty(avd))
                    avd = main.appConfig.Avdeling.ToString();

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblService WHERE (DatoMottat >= '" + DatoFra.ToString("yyy-MM-dd") + "') AND (DatoMottat <= '" + dateArg.ToString("yyy-MM-dd") + "') AND Avdeling = '" + avd + "' ORDER BY DatoMottat ASC");

                if (sqlce.Rows.Count > 0)
                {
                    GraphMaxDay = 0;
                    int antall = 0;
                    for (int o = 0; o < days + 1; o++)
                    {
                        DataRow dtRow = dt.NewRow();
                        dtRow[0] = o;
                        dtRow[1] = dateArg.AddDays(-o);
                        StorageService store = new StorageService();
                        for (int i = 0; i < sqlce.Rows.Count; i++)
                        {
                            if (Convert.ToDateTime(sqlce.Rows[i][4]) == dateArg.AddDays(-o))
                            {
                                var mottat = Convert.ToDateTime(sqlce.Rows[i][4]);
                                var iarbeid = Convert.ToDateTime(sqlce.Rows[i][5]);
                                var ferdig = Convert.ToDateTime(sqlce.Rows[i][6]);
                                var utlevert = Convert.ToDateTime(sqlce.Rows[i][7]);
                                var ordrenr = Convert.ToInt32(sqlce.Rows[i][2]);
                                var sk = sqlce.Rows[i][9].ToString();
                                var stat = status.GetStatusInt(sqlce.Rows[i][8].ToString());
                                var navn = sqlce.Rows[i][3].ToString();
                                bool ferdigbehandlet = false;
                                if (DBNull.Value != sqlce.Rows[i][11])
                                    ferdigbehandlet = Convert.ToBoolean(sqlce.Rows[i][11]);

                                bool egenservice = false;
                                if (navn.Contains(main.appConfig.serviceEgenServiceFilter) && main.appConfig.serviceEgenServiceFilter.Length > 2)
                                    egenservice = true;

                                int tat = 0;
                                if (ferdig.Year > checkValue.Year)
                                    tat = (ferdig - mottat).Days;
                                else
                                    tat = (dbServiceDatoTil - mottat).Days;

                                int tilarbeid = 0;
                                if (iarbeid.Year > checkValue.Year)
                                    tilarbeid = (iarbeid - mottat).Days;
                                else
                                    tilarbeid = (dbServiceDatoTil - mottat).Days;

                                store.servicer.Add(new StorageServieArray(stat, tat, ordrenr, sk, egenservice, iarbeid, tilarbeid, ferdig, utlevert, ferdigbehandlet));
                                antall++;
                            }
                        }
                        store.antall = antall;
                        antall = 0;
                        if (GraphMaxDay < store.antall)
                            GraphMaxDay = store.antall;

                        dtRow[2] = store;
                        dt.Rows.Add(dtRow);
                    }

                    return dt;

                }
                return dt;
            }
            catch (IndexOutOfRangeException ex)
            {
                Log.Unhandled(ex);
                if (MessageBox.Show("Oppdaget en mulig utdatert service database!\nVil du forsøke en oppgradering av databasen?\n\nAdvarsel: Alle importerte servicer vil bli slettet!\n(historisk data vil bli bevart)", "KGSA - Informasjon", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3) == DialogResult.Yes)
                {
                    main.database.tableService.Reset();
                    main.database.tableServiceHistory.Reset();
                    main.database.tableServiceLogg.Reset();

                    dbServiceDatoFra = FormMain.rangeMin;
                    dbServiceDatoTil = FormMain.rangeMin;
                    Log.n("Oppgradering fullført. Importer servicer på nytt!", Color.Green);
                }
                return new DataTable();
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        float maxGraphHistoryValue = 0;

        public DataTable GetServiceHistory(int days)
        {
            try
            {
                var dt = new DataTable();
                dt.Columns.Add("Index", typeof(int));
                dt.Columns.Add("Dato", typeof(DateTime));

                DateTime DatoFra = dbServiceDatoTil.AddDays(-days);
                DateTime checkValue = new DateTime(2001, 1, 1);

                DataTable sqlce = main.database.GetSqlDataTable("SELECT Dato, Totalt, Aktive, Tat, Over14prosent, Over21prosent, Tilarbeid, Avdeling FROM tblServiceHistory");

                if (sqlce.Rows.Count > 0)
                {
                    for (int o = 0; o < days + 1; o++)
                    {
                        DataRow dtRow = dt.NewRow();
                        dtRow[0] = o;
                        dtRow[1] = dbServiceDatoTil.AddDays(-o);
                        dt.Rows.Add(dtRow);
                    }
                    maxGraphHistoryValue = 0;

                    for (int b = 0; b < FormMain.Favoritter.Count; b++)
                    {
                        dt.Columns.Add(FormMain.Favoritter[b], typeof(ServiceHistory));
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            for (int i = 0; i < sqlce.Rows.Count; i++)
                            {
                                if (Convert.ToDateTime(sqlce.Rows[i][0]) == (DateTime)dt.Rows[j][1] && sqlce.Rows[i][7].ToString() == FormMain.Favoritter[b])
                                {
                                    var totalt = Convert.ToInt32(sqlce.Rows[i][1]); // totalt
                                    var aktive = Convert.ToInt32(sqlce.Rows[i][2]); // aktive
                                    var tat = Convert.ToDecimal(sqlce.Rows[i][3]); // tat
                                    var over14prosent = Convert.ToDecimal(sqlce.Rows[i][4]); // over14 prosent
                                    var over21prosent = Convert.ToDecimal(sqlce.Rows[i][5]); // over21 prosent
                                    var tilarbeid = Convert.ToInt32(sqlce.Rows[i][6]); // til arbeid

                                    if (maxGraphHistoryValue < aktive)
                                        maxGraphHistoryValue = aktive;
                                    ServiceHistory store = new ServiceHistory(totalt, aktive, tat, over14prosent, over21prosent, tilarbeid);

                                    dt.Rows[j][b + 2] = store;
                                }
                            }
                        }
                    }

                    return dt;

                }
                return dt;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private DataTable GetServiceList(string statusFilter)
        {
            try
            {
                string filter = "";
                if (statusFilter.Length == 0 || statusFilter == "alle")
                    filter = "";
                else
                    filter = " AND Status = '" + statusFilter + "'";

                DataTable sqlce = main.database.GetSqlDataTable("SELECT * FROM tblService WHERE Avdeling = '" + main.appConfig.Avdeling + "'" + filter + " AND (Status != 'Ferdig, utlevert' AND Status != 'Ferdig/venter utlev')");

                DataTable dt = ReadyTableService();

                if (sqlce.Rows.Count == 0)
                    return dt;

                int totaltDager = 0;
                for (int i = 0; i < sqlce.Rows.Count; i++)
                {
                    DataRow dtRow = dt.NewRow();
                    var date = Convert.ToDateTime(sqlce.Rows[i][4]);

                    dtRow["OrdreNr"] = sqlce.Rows[i][2];
                    dtRow["DatoMottat"] = date;
                    dtRow["Status"] = sqlce.Rows[i][8].ToString();
                    var navn = sqlce.Rows[i][3].ToString();
                    dtRow["Navn"] = navn;
                    if (navn.Contains(main.appConfig.serviceEgenServiceFilter) && main.appConfig.serviceEgenServiceFilter.Length > 2)
                        dtRow["Egenservice"] = true;
                    else
                        dtRow["Egenservice"] = false;
                    dtRow["Selgerkode"] = sqlce.Rows[i][9].ToString();
                    dtRow["Verksted"] = sqlce.Rows[i][10].ToString();
                    dtRow["ServiceID"] = (int)sqlce.Rows[i][0];

                    int dager = (dbServiceDatoTil - date).Days;
                    totaltDager += dager;
                    dtRow["Dager"] = dager;

                    bool behandlet = false;
                    if (DBNull.Value != sqlce.Rows[i][11])
                        behandlet = Convert.ToBoolean(sqlce.Rows[i][11]);
                    dtRow["FerdigBehandlet"] = behandlet;

                    dt.Rows.Add(dtRow);
                }

                DataView dv = dt.DefaultView;
                dv.Sort = "Dager desc";
                DataTable sortedDT = dv.ToTable();

                return sortedDT;

            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        private DataTable GetServiceSpecialList(string filter, string loggFilter)
        {
            try
            {

                string sql = "SELECT tblService.ServiceID, tblService.Ordrenr, tblServiceLogg.Kode, tblServiceLogg.Tekst, tblService.DatoMottat, tblService.Navn, tblService.Selgerkode, tblService.Verksted, tblService.Status, tblService.FerdigBehandlet " +
                    "FROM tblService, tblServiceLogg " +
                    "WHERE tblService.ServiceID = tblServiceLogg.ServiceID AND tblService.Status = 'I arb. eksternt' " +
                    "AND tblService.Avdeling = " + main.appConfig.Avdeling + " " +
                    "ORDER BY tblService.Ordrenr, tblServiceLogg.DatoTid";
                DataTable sqlce = main.database.GetSqlDataTable(sql);

                int c = sqlce.Rows.Count;

                if (c == 0)
                    return null;

                DataTable dt = ReadyTableService();

                int totaltDager = 0;
                int ordreNr = 0;
                for (int i = 0; i < c; i++)
                {
                    if (ordreNr == (int)sqlce.Rows[i][1] && ordreNr != 0)
                    {
                        ordreNr = 0;
                    }
                    else if (ordreNr != 0)
                    {
                        i--;
                        DataRow dtRow = dt.NewRow();
                        var date = Convert.ToDateTime(sqlce.Rows[i][4]);

                        dtRow["OrdreNr"] = sqlce.Rows[i][1];
                        dtRow["DatoMottat"] = date;
                        dtRow["Status"] = sqlce.Rows[i][8].ToString();
                        var navn = sqlce.Rows[i][3].ToString();
                        dtRow["Navn"] = navn;
                        if (navn.Contains(main.appConfig.serviceEgenServiceFilter) && main.appConfig.serviceEgenServiceFilter.Length > 2)
                            dtRow["Egenservice"] = true;
                        else
                            dtRow["Egenservice"] = false;
                        dtRow["Selgerkode"] = sqlce.Rows[i][6].ToString();
                        dtRow["Verksted"] = sqlce.Rows[i][7].ToString();
                        dtRow["ServiceID"] = (int)sqlce.Rows[i][0];

                        int dager = (dbServiceDatoTil - date).Days;
                        totaltDager += dager;
                        dtRow["Dager"] = dager;

                        bool behandlet = false;
                        if (DBNull.Value != sqlce.Rows[i][9])
                            behandlet = Convert.ToBoolean(sqlce.Rows[i][9]);
                        dtRow["FerdigBehandlet"] = behandlet;

                        dt.Rows.Add(dtRow);
                        i++;

                        ordreNr = 0;
                        continue;
                    }

                    if (sqlce.Rows[i][3].ToString().Contains(filter))
                        ordreNr = (int)sqlce.Rows[i][1];
                    else
                        ordreNr = 0;

                }

                DataView dv = dt.DefaultView;
                dv.Sort = "Dager desc";
                DataTable sortedDT = dv.ToTable();

                return sortedDT;

            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return null;
            }
        }

        int importReadErrors = 0;
        public bool Import(string filename, FormProcessing processing, BackgroundWorker bw = null)
        {
            try
            {
                main.processing = processing;
                if (!File.Exists(filename))
                {
                    Log.n("Service Import: Fant ikke CSV service fil eller ble nektet tilgang. (" + filename + ")", Color.Red);
                    return false;
                }

                var engine = new FileHelperEngine(typeof(csvService));
                engine.ErrorManager.ErrorMode = ErrorMode.SaveAndContinue;

                processing.SetProgressStyle = ProgressBarStyle.Marquee;
                engine.SetProgressHandler(new ProgressChangeHandler(ReadProgressCSV));
                var resCSV = engine.ReadFile(filename) as csvService[];

                if (engine.ErrorManager.HasErrors)
                    foreach (ErrorInfo err in engine.ErrorManager.Errors)
                    {
                        importReadErrors++;
                        Log.n("Import: Klarte ikke lese linje " + err.LineNumber + ": " + err.RecordString, Color.Red);
                        Log.d("Exception: " + err.ExceptionInfo.ToString());

                        if (importReadErrors > 100)
                        {
                            Log.n("Feil: CSV er ikke en service eksportering eller filen er skadet. (" + filename + ")", Color.Red);
                            return false;
                        }

                    }
                processing.SetProgressStyle = ProgressBarStyle.Continuous;
                int count = resCSV.Length;

                if (count > 0)
                {
                    DateTime dtFirst = DateTime.MaxValue;
                    DateTime dtLast = DateTime.MinValue;

                    string sqlStrAvd = "";
                    for (int i = 0; i < count; i++)
                    {
                        if (main.appConfig.importSetting != "FullFavoritt")
                            if (!sqlStrAvd.Contains(resCSV[i].Avd.ToString())) // Bygg SQL streng av avdelinger som skal slettes.
                                sqlStrAvd = sqlStrAvd + "Avdeling = '" + resCSV[i].Avd.ToString() + "' OR ";

                        DateTime dtTemp = Convert.ToDateTime(resCSV[i].DatoMottatt.ToString());
                        if (DateTime.Compare(dtTemp, dtFirst) < 0)
                            dtFirst = dtTemp;

                        if (DateTime.Compare(dtTemp, dtLast) > 0)
                            dtLast = dtTemp;
                    }

                    int dager = (dtLast - dtFirst).Days;
                    if (dager < 60)
                    {
                        Log.n("Feil: CSV inneholder for kort periode. Velg minst 60 dager perioder.", Color.Red);
                        return false;
                    }
                    if (dtLast < dbServiceDatoTil)
                    {
                        Log.n("Feil: CSV inneholder bare servicer som er eldre enn eksisterende servicer.", Color.Red);
                        return false;
                    }

                    if (main.appConfig.importSetting == "FullFavoritt")
                        foreach (string avdel in FormMain.Favoritter)
                            sqlStrAvd = sqlStrAvd + "Avdeling = '" + avdel + "' OR ";
                    if (sqlStrAvd.Length > 3) // Remove the last "OR"
                        sqlStrAvd = sqlStrAvd.Remove(sqlStrAvd.Length - 3);


                    var command = new SqlCeCommand("DELETE FROM tblService WHERE (DatoMottat >= '" + dtFirst.ToString("yyy-MM-dd") + "') AND (DatoMottat <= '" + dtLast.ToString("yyy-MM-dd") + "') AND (" + sqlStrAvd + ")", main.connection);
                    var result = command.ExecuteNonQuery();
                    Log.d("Slettet " + result + " servicer.");

                    command = new SqlCeCommand("DELETE FROM tblServiceLogg", main.connection);
                    result = command.ExecuteNonQuery();
                    Log.d("Slettet " + result + " logg oppføringer.");

                    Log.n("Prosesserer " + count.ToString("#,##0") + " service oppføringer.. (" + filename + ")");

                    if (main.appConfig.importSetting == "FullFavoritt")
                        foreach (string avdel in FormMain.Favoritter)
                            sqlStrAvd = sqlStrAvd + "Avdeling = '" + avdel + "' OR ";
                    if (sqlStrAvd.Length > 3) // Remove the last "OR"
                        sqlStrAvd = sqlStrAvd.Remove(sqlStrAvd.Length - 3);

                    int ordreNr = 0;
                    int currentRecord = 0;
                    int currentLog = 0;
                    int id = 0;
                    string sql = "INSERT INTO tblService (Avdeling, Ordrenr, Navn, DatoMottat, DatoIarbeid, DatoFerdig, DatoUtlevert, Status, Selgerkode, Verksted) " +
                        "VALUES (@Avdeling, @Ordrenr, @Navn, @DatoMottat, @DatoIarbeid, @DatoFerdig, @DatoUtlevert, @Status, @Selgerkode, @Verksted)";
                    
                    for (int i = 0; i < count; i++)
                    {
                        if (FormMain.Favoritter.Contains(resCSV[i].Avd.ToString()))
                        {
                            if ((int)resCSV[i].Ordrenr != ordreNr)
                            {
                                ordreNr = (int)resCSV[i].Ordrenr;

                                currentRecord++;
                                processing.SetText = "Lagrer servicer " + currentRecord.ToString("#,##0") + "..";

                                if (bw != null)
                                {
                                    if (bw.WorkerReportsProgress)
                                        bw.ReportProgress(i, new StatusProgress(count, "Prosessere servicer oppføringer..", 0, 100));
                                    if (bw.CancellationPending)
                                    {
                                        Log.n("Service importering avbrutt av bruker!", Color.Red);
                                        return false;
                                    }
                                }

                                using (SqlCeCommand cmd = new SqlCeCommand(sql, main.connection))
                                {
                                    cmd.Parameters.AddWithValue("@Avdeling", (int)resCSV[i].Avd);
                                    cmd.Parameters.AddWithValue("@Ordrenr", (int)resCSV[i].Ordrenr);
                                    cmd.Parameters.AddWithValue("@Navn", resCSV[i].Navn);
                                    cmd.Parameters.Add("@DatoMottat", SqlDbType.DateTime).Value = Convert.ToDateTime(resCSV[i].DatoMottatt);
                                    var date = Convert.ToDateTime(resCSV[i].DatoIarbeid);
                                    if (date > FormMain.rangeMin)
                                        cmd.Parameters.Add("@DatoIarbeid", SqlDbType.DateTime).Value = date;
                                    else
                                        cmd.Parameters.Add("@DatoIarbeid", SqlDbType.DateTime).Value = FormMain.rangeMin;
                                    date = Convert.ToDateTime(resCSV[i].DatoFerdig);
                                    if (date > FormMain.rangeMin)
                                        cmd.Parameters.Add("@DatoFerdig", SqlDbType.DateTime).Value = date;
                                    else
                                        cmd.Parameters.Add("@DatoFerdig", SqlDbType.DateTime).Value = FormMain.rangeMin;
                                    date = Convert.ToDateTime(resCSV[i].DatoUtlevert);
                                    if (date > FormMain.rangeMin)
                                        cmd.Parameters.Add("@DatoUtlevert", SqlDbType.DateTime).Value = date;
                                    else
                                        cmd.Parameters.Add("@DatoUtlevert", SqlDbType.DateTime).Value = FormMain.rangeMin;
                                    cmd.Parameters.AddWithValue("@Status", resCSV[i].Status);
                                    cmd.Parameters.AddWithValue("@Selgerkode", resCSV[i].Selgerkode);
                                    cmd.Parameters.AddWithValue("@Verksted", resCSV[i].Verksted);
                                    cmd.CommandType = System.Data.CommandType.Text;
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "SELECT @@IDENTITY";
                                    id = Convert.ToInt32(cmd.ExecuteScalar());
                                }
                                currentLog = 1;
                                ImportAddLog(Convert.ToDateTime(resCSV[i].LoggDato), Convert.ToDateTime(resCSV[i].LoggTid), resCSV[i].LoggKode, resCSV[i].LoggTekst, id, currentRecord, currentLog, main.connection);
                            }
                            else
                            {
                                currentLog++;
                                ImportAddLog(Convert.ToDateTime(resCSV[i].LoggDato), Convert.ToDateTime(resCSV[i].LoggTid), resCSV[i].LoggKode, resCSV[i].LoggTekst, id, currentRecord, currentLog, main.connection);
                                continue;
                            }
                        }
                    }
                }
                return true;
            }
            catch (IOException ex)
            {
                Log.n("CSV var låst for lesing. Forleng ventetid i makro hvis overføringen ikke ble ferdig i tide.", Color.Red);
                Log.Unhandled(ex);
                return false;
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil ved importering av servicer", ex);
                errorMsg.ShowDialog();
                return false;
            }
        }

        private void ImportAddLog(DateTime argDate, DateTime argTime, string argKode, string argTekst, int id, int currentRecord, int currentLog, SqlCeConnection con)
        {

            try
            {
                string command = "INSERT INTO tblServiceLogg (ServiceID, DatoTid, Kode, Tekst) " +
            "VALUES (@ServiceID, @DatoTid, @Kode, @Tekst)";

                currentLog++;
                main.processing.SetText = "Lagrer servicer " + currentRecord.ToString("#,##0") + ".. (Logg " + currentLog + ")";

                using (SqlCeCommand cmd = new SqlCeCommand(command, con))
                {
                    cmd.Parameters.AddWithValue("@ServiceID", id);

                    TimeSpan timeSpan = new TimeSpan(argTime.Hour, argTime.Minute, 0);
                    DateTime combined = argDate.Add(timeSpan);

                    if (combined > FormMain.rangeMin)
                        cmd.Parameters.Add("@DatoTid", SqlDbType.DateTime).Value = combined;
                    else
                        cmd.Parameters.Add("@DatoTid", SqlDbType.DateTime).Value = FormMain.rangeMin;
                    cmd.Parameters.AddWithValue("@Kode", argKode);
                    cmd.Parameters.AddWithValue("@Tekst", argTekst);

                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }

        }

        public void NullstillBehandlet()
        {
            try
            {
                if (dbServiceDatoFra == dbServiceDatoTil)
                {
                    Log.n("Service databasen er tom!");
                    return;
                }

                string strDrop = "ALTER TABLE tblService DROP COLUMN [FerdigBehandlet]";
                string strAdd = "ALTER TABLE tblService ADD COLUMN [FerdigBehandlet] bit NULL";

                Log.n("Nullstiller markering..");

                var command = new SqlCeCommand(strDrop, main.connection);
                command.ExecuteNonQuery();
                command = new SqlCeCommand(strAdd, main.connection);
                command.ExecuteNonQuery();

                Log.n("Markering nullstilt.", Color.Green);
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under reset av markeringer.", ex);
                errorMsg.ShowDialog();
            }
        }

        private void ReadProgressCSV(ProgressEventArgs e)
        {
            if (e.ProgressCurrent % 831 == 0)
                main.processing.SetText = "Leser CSV: " + e.ProgressCurrent.ToString("#,##0") + "..";
        }
    }

    public class ServiceStat
    {
        public DateTime oppdatert { get; set; }
        public int avdeling { get; set; }
        public int aktive { get; set; }
        public int totalt { get; set; }
        public int ferdig { get; set; }
        public int over14aktiv { get; set; }
        public int over21aktiv { get; set; }
        public int over30aktiv { get; set; }
        public decimal tat60 { get; set; }
        public float over14aktivprosent { get; set; }
        public float over21aktivprosent { get; set; }
        public float over30aktivprosent { get; set; }
        public decimal tilarbeid60 { get; set; }
        public int venter { get; set; }
        public int over14 { get; set; }
        public float over14prosent { get; set; }
        public int over21 { get; set; }
        public float over21prosent { get; set; }
        public int over30 { get; set; }
        public float over30prosent { get; set; }
        public ServiceStat()
        {
            aktive = 0;
            totalt = 0;
            ferdig = 0;
            over14aktiv = 0;
            over30aktiv = 0;
            tat60 = 0;
            over14aktivprosent = 0;
            over21aktivprosent = 0;
            tilarbeid60 = 0;
            venter = 0;
            over14 = 0;
            over21 = 0;
            over30 = 0;
            over14prosent = 0;
            over21prosent = 0;
            over30prosent = 0;
        }
    }
}
