using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

namespace KGSA
{
    public class KgsaServer
    {
        public FormMain main;
        public WebServer ws;

        public KgsaServer(FormMain form)
        {
            this.main = form;
        }

        public void StartWebserver()
        {
            try
            {
                if (main.appConfig.webserverEnabled && main.appConfig.webserverPort >= 1024 && main.appConfig.webserverPort <= 65535 && !String.IsNullOrEmpty(main.appConfig.webserverHost))
                {
                    if (ws != null)
                        if (ws.IsOnline())
                            ws.Stop();

                    //AddAddress(); // Åpne brannmuren
                    //System.Threading.Thread.Sleep(1000);

                    ws = new WebServer(ResponseDummy, "http://" + main.appConfig.webserverHost + ":" + main.appConfig.webserverPort + "/");
                    ws.Settings(main.appConfig, this);
                    ws.Start();
                    ws.Run();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        //private void AddAddress()
        //{
        //    try
        //    {
        //        string args = "netsh http add urlacl http://+:" + main.appConfig.webserverPort + "/ user=Everyone listen=on";

        //        ProcessStartInfo psi = new ProcessStartInfo("netsh", args);
        //        psi.Verb = "runas";
        //        psi.CreateNoWindow = true;
        //        psi.WindowStyle = ProcessWindowStyle.Hidden;
        //        psi.UseShellExecute = true;

        //        Process.Start(psi).WaitForExit();
        //    }
        //    catch(Exception ex)
        //    {
        //        Logg.Unhandled(ex);
        //    }
        //}

        public void ProcessRequest(HttpListenerContext ctx)
        {
            string chkstr = ctx.Request.RawUrl;
            Logg.Debug("WebRequest: " + chkstr);

            if (chkstr.Length > 0 && chkstr.Length < 100 && !chkstr.Contains(".."))
            {
                string file = HttpUtility.UrlDecode(chkstr);

                if (File.Exists(FormMain.settingsPath + file) && file.Length > 3) // Bruker spør etter fil som finnes..
                {
                    if (file.EndsWith(".html"))
                    {
                        string page = File.ReadAllText(FormMain.settingsPath + file);
                        page = page.Replace("    /*padding-top: 23px;*/", "    padding-top: 23px;");
                        page = page.Replace("<!-- header here -->", InsertHeader(file, DateTime.Now, ""));
                        page = page.Replace("<!-- jsWeb here -->", Resources.jsWeb);


                        byte[] chkbuf = Encoding.UTF8.GetBytes(page);
                        ctx.Response.ContentLength64 = chkbuf.Length;
                        ctx.Response.OutputStream.Write(chkbuf, 0, chkbuf.Length);
                        ctx.Response.ContentEncoding = System.Text.Encoding.UTF8;
                    }
                    else if (file.EndsWith(".js") || file.EndsWith(".csv"))
                    {
                        string page = File.ReadAllText(FormMain.settingsPath + file);
                        byte[] chkbuf = Encoding.UTF8.GetBytes(page);
                        ctx.Response.ContentLength64 = chkbuf.Length;
                        ctx.Response.OutputStream.Write(chkbuf, 0, chkbuf.Length);
                        ctx.Response.ContentEncoding = System.Text.Encoding.UTF8;
                    }
                    else if (file.EndsWith(".ico") || file.EndsWith(".png") || file.EndsWith(".pdf"))// andre filer (bilder osv)
                    {
                        var byters = File.ReadAllBytes(FormMain.settingsPath + chkstr);
                        ctx.Response.ContentLength64 = byters.Length;
                        ctx.Response.OutputStream.Write(byters, 0, byters.Length);
                    }
                    else // Filen finnes men er ikke blandt de gyldige filtypene, vi sender tilbake en 404 side..
                    {
                        ctx.Response.StatusCode = 404;
                        string rstr = WebPage_FileNotFound();
                        byte[] chkbuf = Encoding.UTF8.GetBytes(rstr);
                        ctx.Response.ContentLength64 = chkbuf.Length;
                        ctx.Response.OutputStream.Write(chkbuf, 0, chkbuf.Length);
                        ctx.Response.ContentEncoding = System.Text.Encoding.UTF8;
                    }

                }
                else // Ok, ingen fil eller filen eksisterer ikke. Sendes videre til dynamisk side generator..
                {
                    string rstr = ResponseServer(ctx);
                    byte[] buf = Encoding.UTF8.GetBytes(rstr);
                    ctx.Response.ContentLength64 = buf.Length;
                    ctx.Response.OutputStream.Write(buf, 0, buf.Length);
                    ctx.Response.ContentEncoding = System.Text.Encoding.UTF8;
                }
            }
        }

        private string InsertHeader(string file, DateTime date, string katArg)
        {
            List<string> doc = new List<string>() { };

            doc.Add("<header class='webmenu'><section>");
            doc.Add("<a href=\"/\">Forsiden</a> | ");
            if (main.appConfig.importSetting.StartsWith("Full"))
            {
                if (File.Exists(FormMain.htmlRankingToppselgere))
                    doc.Add("<a href=\"rankingToppselgere.html\">Toppselgere</a> | ");
                if (File.Exists(FormMain.htmlServiceOversikt))
                    doc.Add("<a href=\"rankingOversikt.html\">Oversikt</a> | ");
            }
            if (File.Exists(FormMain.htmlRankingButikk))
                doc.Add("<a href=\"rankingButikk.html\">Butikk</a> | ");
            if (File.Exists(FormMain.htmlRankingKnowHow))
                doc.Add("<a href=\"rankingKnowHow.html\">KnowHow</a> | ");
            if (File.Exists(FormMain.htmlRankingData))
                doc.Add("<a href=\"rankingData.html\">Data</a> | ");
            if (File.Exists(FormMain.htmlRankingAudioVideo))
                doc.Add("<a href=\"rankingAudioVideo.html\">Lyd & Bilde</a> | ");
            if (File.Exists(FormMain.htmlRankingTele))
                doc.Add("<a href=\"rankingTele.html\">Tele</a> | ");
            if (File.Exists(FormMain.htmlRankingLister))
                doc.Add("<a href=\"rankingLister.html\">Lister</a> |");
            if (File.Exists(FormMain.htmlRankingVinn))
                doc.Add("<a href=\"rankingVinnprodukter.html\">Vinnprodukter</a> |");
            if (!String.IsNullOrEmpty(file))
                doc.Add("<a href=\"/content?p=pdf&action=" + file + "\">Last ned PDF</a> |");

            if (date.Date != DateTime.Now.Date) // Siden har mulighet for å hente ut andre datoer
            {
                doc.Add("Dag: <form  action='' method='get'><input type='hidden' name='p' value='" + katArg + "'><select name='day' id='selectDay'>");
                for (int i = 1; i < 32; i++)
                {
                    if (i == date.Day)
                        doc.Add("<option value='" + i + "' selected>" + i + "</option>");
                    else
                        doc.Add("<option value='" + i + "'>" + i + "</option>");
                }
                doc.Add("</select>");

                doc.Add("Måned: <select name='month' id='selectMonth'>");
                for (int i = 1; i < 13; i++)
                {
                    if (i == date.Month)
                        doc.Add("<option value='" + i + "' selected>" + new DateTime(2001, i, 1).ToString("MMM", FormMain.norway) + "</option>");
                    else
                        doc.Add("<option value='" + i + "'>" + new DateTime(2001, i, 1).ToString("MMM", FormMain.norway) + "</option>");
                }
                doc.Add("</select>");

                doc.Add("År: <select name='year' id='selectYear'>");
                for (int i = main.appConfig.dbFrom.Year; i <= main.appConfig.dbTo.Year ; i++)
                {
                    if (i == date.Year)
                        doc.Add("<option value='" + i + "' selected>" + i + "</option>");
                    else
                        doc.Add("<option value='" + i + "'>" + i + "</option>");
                }
                doc.Add("</select> <button type='submit'>Gå til dato</button></form>");
                if (date.Date != main.appConfig.dbTo.Date)
                {
                    doc.Add("<form action='' method='get'><input type='hidden' name='p' value='" 
                        + katArg + "'><input type='hidden' name='day' value='" + main.appConfig.dbTo.Day
                        + "'><input type='hidden' name='month' value='" + main.appConfig.dbTo.Month
                        + "'><input type='hidden' name='year' value='" + main.appConfig.dbTo.Year + "'>");
                    doc.Add("<button type='submit'>Gå til siste (" + main.appConfig.dbTo.ToShortDateString() + ")</button></form>");
                }
            }

            doc.Add("</section></header>");

            return string.Join(Environment.NewLine, doc.ToArray());
        }

        public string ResponseDummy(HttpListenerRequest request)
        {
            return "";
        }

        public string ResponseServer(HttpListenerContext ctx )
        {
            HttpListenerRequest request = ctx.Request;

            string page = "", action = "", date = "", day = "", month = "", year = "";
            int id = 0;
            string url = System.Web.HttpUtility.UrlDecode(request.RawUrl);

            if (url.Contains("?")) // Parse queries
            {
                var queries = HttpUtility.ParseQueryString(url.Substring(url.IndexOf('?')));
                if (!queries["p"].IsNullOrWhiteSpace())
                    page = queries["p"];
                if (!queries["action"].IsNullOrWhiteSpace())
                    action = queries["action"];
                if (!queries["date"].IsNullOrWhiteSpace())
                    date = queries["date"];
                if (!queries["id"].IsNullOrWhiteSpace())
                    id = Convert.ToInt32(queries["id"]);

                if (!queries["day"].IsNullOrWhiteSpace())
                    day = queries["day"];
                if (!queries["month"].IsNullOrWhiteSpace())
                    month = queries["month"];
                if (!queries["year"].IsNullOrWhiteSpace())
                    year = queries["year"];
            }

            if (url == "/") // Home page
                return WebPage_HomePage(request);

            if (url.StartsWith("/settings"))
            {
                if (String.IsNullOrEmpty(page) && String.IsNullOrEmpty(action))
                    return WebPage_Settings(request);

                if (page == "save")
                {
                    if (String.IsNullOrEmpty(action))
                        return WebPage_Settings_Save(request);
                }

                if (page == "email")
                {
                    if (String.IsNullOrEmpty(action))
                        return WebPage_Settings_Email(request);
                    else if (action == "add")
                        return WebPage_Settings_EmailAdd(request);
                    else if (action == "remove" && id != 0)
                        return WebPage_Settings_EmailRemove(request, id);
                }

                if (page == "vinn")
                {
                    if (String.IsNullOrEmpty(action))
                        return WebPage_Settings_Vinn(request);
                    else if (action == "add")
                        return WebPage_Settings_VinnAdd(request);
                    else if (action == "remove" && id != 0)
                        return WebPage_Settings_VinnRemove(request, id);
                }

                if (page == "process")
                {
                    if (action == "update" || action == "autoranking" || action == "importstore" || action == "importservice")
                        return WebPage_StartProcess(request, action);
                }

                if (page == "budget")
                {
                    if (String.IsNullOrEmpty(action))
                        return WebPage_Settings_Budget(request);
                }
            }

            if (url.StartsWith("/status"))
            {
                if (String.IsNullOrEmpty(page) && String.IsNullOrEmpty(action))
                    return WebPage_Status(request);
            }

            if (url.StartsWith("/content"))
            {
                if (page == "pdf" && !String.IsNullOrEmpty(action))
                    return WebPage_CreatePdf(request, action);
            }

            if (url.StartsWith("/ranking"))
            {
                if (page == "Data" || page == "AudioVideo" || page == "Tele")
                    return WebPage_Ranking(request, page, day, month, year);
            }

            ctx.Response.StatusCode = 404;
            return WebPage_FileNotFound();
        }

        private string WebPage_CreatePdf(HttpListenerRequest request, string html)
        {
            try
            {
                List<string> doc = new List<string>() { };

                string pdfFile = "download.pdf";

                if (html.Length > 4)
                {
                    pdfFile = Path.GetFileName(html).ToLower();
                    pdfFile = pdfFile.Replace(".html", ".pdf");
                    pdfFile = @"Temp\" + pdfFile;
                }

                bool generated = false;

                if (html.EndsWith(".html"))
                {
                    if (File.Exists(FormMain.settingsPath + @"\" + html))
                        generated = CreatePdf(FormMain.settingsPath + @"\" + html, pdfFile);
                    else
                        Logg.WebUser("Bruker har bedt om PDF av fil som ikke finnes. Fil: " + html, request);
                }
                else
                    Logg.WebUser("Webbruker har bedt om PDF av ugyldig filtype. Fil: " + html, request);

                WebMenuStart(doc, "PDF");

                if (generated)
                {
                    Logg.WebUser("Lager PDF (" + pdfFile + ")", request);
                    doc.Add("Last ned PDF: <a href='" + pdfFile + "'>Download</a>");
                }
                else
                {
                    doc.Add("Feil oppstod under generering av PDF eller server var opptatt.");

                    doc.Add("<br><br>Logg:<br>");
                    foreach (string loggLine in FormMain.loggCacheMessages)
                        doc.Add(loggLine + "<br>");
                }

                WebMenuEnd(doc);


                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private bool CreatePdf(string source, string destination)
        {
            try
            {

                if (!File.Exists(source))
                {
                    Logg.Log("PDF: Fant ikke fil " + source, Color.Red);
                    return false;
                }
                string argSource = " \"" + source + "\" ";

                string options = " -B 20 -L 7 -R 7 -T 7 --zoom " + main.appConfig.pdfZoom + " ";
                if (main.appConfig.pdfLandscape)
                    options += "-O landscape ";

                Logg.Debug("PDF simple generator argument: " + options + argSource + destination);

                var wkhtmltopdf = new ProcessStartInfo();
                wkhtmltopdf.WindowStyle = ProcessWindowStyle.Hidden;
                wkhtmltopdf.FileName = FormMain.filePDFwkhtmltopdf;
                wkhtmltopdf.Arguments = options + argSource + destination;
                wkhtmltopdf.WorkingDirectory = FormMain.settingsPath;
                wkhtmltopdf.CreateNoWindow = true;
                wkhtmltopdf.UseShellExecute = false;

                Process D = Process.Start(wkhtmltopdf);

                D.WaitForExit(20000);

                if (!D.HasExited)
                {
                    Logg.Log("Error: PDF generatoren ble tidsavbrutt.", Color.Red);
                    return false;
                }

                int result = D.ExitCode;
                if (result != 0)
                {
                    Logg.Log("Error: PDF generator returnerte med feilkode " + result, Color.Red);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        private string WebPage_Status(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                if (!main.IsBusy())
                    WebMenuStart(doc, "Status");
                else
                    WebMenuStart(doc, "Status", "", true);

                if (main.IsBusy())
                    doc.Add("KGSA er opptatt med oppgaver.");
                else
                    doc.Add("<span style='color:green;'>KGSA er klar.</span>");

                doc.Add("<br><br>Logg:<br>");
                foreach (string loggLine in FormMain.loggCacheMessages)
                    doc.Add(loggLine + "<br>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_StartProcess(HttpListenerRequest request, string call)
        {
            try
            {
                if (call == "importranking")
                    Logg.Log("Webbruker har bedt om importering av transaksjoner.", Color.DarkGoldenrod);
                else if (call == "importstore")
                    Logg.Log("Webbruker har bedt om importering av lager.", Color.DarkGoldenrod);
                else if (call == "importservice")
                    Logg.Log("Webbruker har bedt om importering av servicer.", Color.DarkGoldenrod);
                else if (call == "update")
                    Logg.Log("Webbruker har bedt om oppdatering av ranking.", Color.DarkGoldenrod);
                else if (call == "autoranking")
                    Logg.Log("Webbruker har bedt om importering og utsending av ranking.", Color.DarkGoldenrod);

                main.WebStartProcess(call); // Start prosessen

                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Status", "../status");

                doc.Add("Kommando mottat.");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_HomePage(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Forside");

                doc.Add("<table><tr>");
                doc.Add("<td width=300>");
                if (!main.EmptyDatabase())
                {
                    doc.Add("Ranking sider:<br>");
                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        if (File.Exists(FormMain.htmlRankingToppselgere))
                            doc.Add("<a href=\"rankingToppselgere.html\">Toppselgere</a><br>");
                        if (File.Exists(FormMain.htmlRankingOversikt))
                            doc.Add("<a href=\"rankingOversikt.html\">Oversikt</a><br>");
                    }
                    if (File.Exists(FormMain.htmlRankingButikk))
                        doc.Add("<a href=\"rankingButikk.html\">Butikk</a><br>");
                    if (File.Exists(FormMain.htmlRankingKnowHow))
                        doc.Add("<a href=\"rankingKnowHow.html\">KnowHow</a><br>");
                    if (File.Exists(FormMain.htmlRankingData))
                        doc.Add("<a href=\"rankingData.html\">Data</a> &nbsp; (<a href='/ranking?p=Data'>Dynamisk</a>)<br>");
                    if (File.Exists(FormMain.htmlRankingAudioVideo))
                        doc.Add("<a href=\"rankingAudioVideo.html\">AudioVideo</a> &nbsp; (<a href='/ranking?p=AudioVideo'>Dynamisk</a>)<br>");
                    if (File.Exists(FormMain.htmlRankingTele))
                        doc.Add("<a href=\"rankingTele.html\">Tele</a> &nbsp; (<a href='/ranking?p=Tele'>Dynamisk</a>)<br>");
                    if (File.Exists(FormMain.htmlRankingLister))
                        doc.Add("<a href=\"rankingLister.html\">Lister</a><br>");
                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        if (File.Exists(FormMain.htmlRankingVinn))
                            doc.Add("<a href=\"rankingVinnprodukter.html\">Vinnprodukter</a><br>");
                    }
                    doc.Add("Oppdatert: " + main.appConfig.dbTo.ToString("dddd d. MMMM yyyy", FormMain.norway) + "<br>");

                    if (main.appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<br>Budsjett sider:<br>");
                        if (File.Exists(FormMain.htmlBudgetMda) && main.appConfig.budgetShowMda)
                            doc.Add("<a href=\"budsjettMda.html\">MDA</a><br>");
                        if (File.Exists(FormMain.htmlBudgetAudioVideo) && main.appConfig.budgetShowAudioVideo)
                            doc.Add("<a href=\"budsjettAudioVideo.html\">Lyd og Bilde</a><br>");
                        if (File.Exists(FormMain.htmlBudgetSda) && main.appConfig.budgetShowSda)
                            doc.Add("<a href=\"budsjettSda.html\">SDA</a><br>");
                        if (File.Exists(FormMain.htmlBudgetTele) && main.appConfig.budgetShowTele)
                            doc.Add("<a href=\"budsjettTele.html\">Tele</a><br>");
                        if (File.Exists(FormMain.htmlBudgetData) && main.appConfig.budgetShowData)
                            doc.Add("<a href=\"budsjettData.html\">Data</a><br>");
                        if (File.Exists(FormMain.htmlBudgetCross) && main.appConfig.budgetShowCross)
                            doc.Add("<a href=\"budsjettCross.html\">Cross</a><br>");
                        if (File.Exists(FormMain.htmlBudgetKasse) && main.appConfig.budgetShowKasse)
                            doc.Add("<a href=\"budsjettKasse.html\">Kasse</a><br>");
                        if (File.Exists(FormMain.htmlBudgetAftersales) && main.appConfig.budgetShowAftersales)
                            doc.Add("<a href=\"budsjettAftersales.html\">Aftersales</a><br>");
                        if (File.Exists(FormMain.htmlBudgetMdasda) && main.appConfig.budgetShowMdasda)
                            doc.Add("<a href=\"budsjettMdaSda.html\">MDA + SDA</a><br>");
                    }
                }
                else
                    doc.Add("Databasen er tom!");
                doc.Add("</td>");

                doc.Add("<td width=300>");
                if (!main.EmptyStoreDatabase())
                {
                    doc.Add("Lager sider:<br>");
                    if (File.Exists(FormMain.htmlStoreObsolete))
                        doc.Add("<a href=\"storeLagerstatus.html\">Lagerstatus</a><br>");
                    if (File.Exists(FormMain.htmlStoreObsoleteList))
                        doc.Add("<a href=\"storeObsoleteList.html\">Ukurant liste</a><br>");
                    if (File.Exists(FormMain.htmlStoreObsoleteImports))
                        doc.Add("<a href=\"storeObsoleteImports.html\">Import liste</a><br>");
                    doc.Add("Oppdatert: " + main.service.dbServiceDatoTil.ToString("dddd d. MMMM yyyy", FormMain.norway) + "<br><br>");
                }
                if (main.service.dbServiceDatoFra.Date != main.service.dbServiceDatoTil)
                {
                    doc.Add("Service sider:<br>");
                    if (File.Exists(FormMain.htmlServiceOversikt))
                        doc.Add("<a href=\"serviceOversikt.html\">Service oversikt</a><br>");
                    doc.Add("Oppdatert: " + main.service.dbServiceDatoTil.ToString("dddd d. MMMM yyyy", FormMain.norway) + "<br>");
                }
                doc.Add("</td></tr></table>");

                doc.Add("<br><br><a href='settings'>Innstillinger</a>");
                doc.Add("<br><br><a href='status'>Status</a>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Ranking(HttpListenerRequest request, string katArg, string day, string month, string year)
        {
            string dateStr = "";
            if (!day.IsNullOrWhiteSpace())
                if (day.Length == 2)
                    dateStr += day;
                else
                    dateStr += "0" + day;
            if (!month.IsNullOrWhiteSpace())
                if (month.Length == 2)
                    dateStr += "." + month;
                else
                    dateStr += ".0" + month;
            if (!year.IsNullOrWhiteSpace())
                dateStr += "." + year;

            DateTime date;

            if (dateStr.Length != 10)
                date = main.appConfig.dbTo;
            else
            {
                if (!DateTime.TryParseExact(dateStr, "dd.MM.yyyy", FormMain.norway, DateTimeStyles.None, out date))
                    date = DateTime.Now;
            }
            Logg.WebUser("Bruker har forespurte om ranking [" + katArg + "] fra dato: " + date.ToShortDateString(), request);

            List<string> doc = new List<string>() { };

            if (date.Date == DateTime.Now.Date)
            {
                main.GetHtmlStart(doc, false, "Ranking: " + katArg);
                doc.Add("<span class='xHeader'>Ugyldig dato angitt.</span><br><span class='xHeader'>Dato: " + dateStr + "</span>");
                doc.Add(Resources.htmlEnd);
            }
            else if (date > main.appConfig.dbTo || date < main.appConfig.dbFrom)
            {
                main.GetHtmlStart(doc, false, "Ranking: " + katArg);
                doc.Add("<span class='xHeader'>Fant ikke transaksjoner for denne datoen</span><br><span class='xHeader'>Dato: " + date.ToShortDateString() + "</span>");
                doc.Add(Resources.htmlEnd);
            }
            else if (katArg == "Data")
                main.BuildWebRankingData(doc, date);
            else if (katArg == "AudioVideo")
                main.BuildWebRankingAudioVideo(doc, date);
            else if (katArg == "Tele")
                main.BuildWebRankingTele(doc, date);
            else
            {
                main.GetHtmlStart(doc, false, "Ranking: " + katArg);
                doc.Add("<span class='xHeader'>Fant ikke den forespurte ranking siden.</span><br><span class='xHeader'>Dato: " + date.ToShortDateString() + "</span>");
                doc.Add(Resources.htmlEnd);
            }

            string page = string.Join(Environment.NewLine, doc.ToArray());

            page = page.Replace("    /*padding-top: 23px;*/", "    padding-top: 23px;");
            page = page.Replace("<!-- header here -->", InsertHeader("", date, katArg));
            page = page.Replace("<!-- jsWeb here -->", Resources.jsWeb);

            return page;
        }

        private string WebPage_Settings_Vinn(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: Vinnprodukter");

                List<VinnproduktItem> list = main.vinnprodukt.GetList();
                if (list.Count > 0)
                {
                    doc.Add("<div class='Table' style='width:650px'><div class='Title'>Vinnprodukter</div>");
                    doc.Add("<div class='Heading'>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Varekode</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Kategori</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Varegruppe</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Poeng</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Dato</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Slett</p>");
                    doc.Add("</div>");
                    doc.Add("</div>");

                    foreach(VinnproduktItem item in list)
                    {
                        doc.Add("<div class='Row'>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + item.varekode + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + item.kategori + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell' style='text-align:center;'>");
                        doc.Add("<p>" + item.poeng + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + item.dato.ToShortDateString() + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell' style='text-align:center;'>");
                        doc.Add("<p><a href='?p=vinn&action=remove&id=" + item.id + "' onclick='return confirmRemove()'>Slett</a></p>");
                        doc.Add("</div>");
                        doc.Add("</div>");
                    }
                    doc.Add("</div>");
                }
                else
                    doc.Add("<span><b>Listen er tom.</b></span>");

                doc.Add("<br><br><form name='form1' method='post' action='/settings?p=vinn&action=add'>");
                doc.Add("<div class='Table' style='width:650px'><div class='Title'>Legg til</div>");
                doc.Add("<div class='Heading'>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Varekode</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Kategori</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Varegruppe</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Poeng</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Action</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='text' name='vinnVarekode' id='vinnVarekode'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><select name='vinnKategori' id='vinnKategori'>");
                doc.Add("<option value='Alle'>Alle</option><option value='MDA'>MDA</option><option value='AudioVideo'>AudioVideo</option><option value='SDA'>SDA</option><option value='Tele'>Tele</option><option value='Data'>Data</option><option value='Cross'>Cross</option><option value='Kasse'>Kasse</option><option value='Aftersales'>Aftersales</option>");
                doc.Add("</select></p></div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='text' name='vinnVaregruppe' id='vinnVaregruppe' style='width:75px;'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='text' name='vinnPoeng' id='vinnPoeng'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='submit' name='Lagre' id='submit' value='Lagre'></p>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("</form>");

                // Setting table top
                doc.Add("<br><br><form id='Settings' name='Settings' method='post' action='/settings?p=save'>");
                doc.Add("<div class='Table Settings'><div class='Title' style='color:white;'>Andre tjenester:</div>");

                doc.Add("<div class='Heading'>");
                doc.Add("<div class='Cell' style='border:none;'>Tjeneste</div>");
                doc.Add("<div class='Cell' style='border:none;'>Poeng</div>");
                doc.Add("<div class='Cell' style='border:none;'>&nbsp;</div>");
                doc.Add("</div>");

                doc.AddRange(CreateRowSettingFor("vinnTjenFinansPoeng", "Finans:", ""));
                doc.AddRange(CreateRowSettingFor("vinnTjenTaPoeng", "Trygghetsavtale::", ""));
                doc.AddRange(CreateRowSettingFor("vinnTjenStromPoeng", "Norges Energi:", ""));

                // Settings table bottom
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;'>&nbsp;</div>");
                doc.Add("<div class='Cell' style='border:none;'>&nbsp;</div>");
                doc.Add("<div class='Cell' style='border:none;padding-bottom:5px;padding-top:5px;'><div style='text-align:right;'><input type='submit' name='Lagre' id='submit' value='Lagre'></div></div>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("<input type='hidden' name='returnPage' id='returnPage' value='settings?p=vinn'>");
                doc.Add("</form>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_VinnAdd(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: Vinnprodukter", "?p=vinn");

                Dictionary<string, string> postParams = ParseParameters(request);

                if (postParams.ContainsKey("vinnVarekode") && postParams.ContainsKey("vinnKategori") && postParams.ContainsKey("vinnPoeng"))
                {
                    string varekode = postParams["vinnVarekode"];
                    string kategori = postParams["vinnKategori"];

                    int varegruppe = 0;
                    if (!int.TryParse(postParams["vinnVaregruppe"], out varegruppe))
                        varegruppe = 0;

                    decimal poeng = 0;
                    if (!decimal.TryParse(postParams["vinnPoeng"], out poeng))
                        poeng = 0;

                    if (main.vinnprodukt.Add(varekode.ToUpper(), kategori, poeng, DateTime.Now.AddYears(1), DateTime.Now.AddYears(-1)))
                    {
                        Logg.WebUser("Bruker har lagt til varekoden " + varekode + " til kategorien " + kategori + " som vinnprodukt.", request);
                        doc.Add("<span style='color:green;'><b>Varekoden " + varekode + " lagt til kategorien " + kategori + ".</b></span>");
                    }
                    else
                        doc.Add("<span style='color:orange;'><b>Kunne ikke lagre varekode. Finnes kanskje fra før?</b></span>");
                }
                else
                    doc.Add("<span style='color:orange;'><b>Manglet parametere.</b></span>");

                main.vinnprodukt.Update();

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_VinnRemove(HttpListenerRequest request, int id)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: Vinnprodukter", "?p=vinn");

                if (main.vinnprodukt.Remove(id.ToString()))
                {
                    Logg.Log("Web: Webbruker slettet id " + id + " fra vinnprodukter.", Color.DarkGoldenrod);
                    doc.Add("<span style='color:greem;'><b>Varkoden med id " + id + " ble slette fra vinnprodukter listen.</b></span>");
                }
                else
                    doc.Add("<span style='color:orange;'><b>Kunne ikke slette varekoden.</b></span>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_EmailAdd(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: E-post", "?p=email");

                Dictionary<string, string> postParams = ParseParameters(request);

                if (postParams.ContainsKey("emailName") && postParams.ContainsKey("emailAddress") && postParams.ContainsKey("emailGroup"))
                {
                    string name = postParams["emailName"];
                    string address = System.Web.HttpUtility.UrlDecode(postParams["emailAddress"]);
                    string group = postParams["emailGroup"];
                    bool quick = false;
                    if (postParams.ContainsKey("emailQuick"))
                    {
                        string str = postParams["emailQuick"].ToLower();
                        if (str == "on" || str == "true")
                            quick = true;
                    }

                    KgsaEmail email = new KgsaEmail(main);
                    if (email.Add(name, address, group, quick))
                    {
                        Logg.Log("Web: Webbruker la til E-post adressen " + address + " i gruppen " + group + " til adresseboken.", Color.DarkGoldenrod);
                        doc.Add("<span style='color:green;'><b>E-post adressen " + address + " lagt til gruppen " + group + ".</b></span>");
                    }
                    else
                        doc.Add("<span style='color:orange;'><b>Kunne ikke lagre e-posten. Finnes kanskje fra før.</b></span>");
                }
                else
                    doc.Add("<span style='color:orange;'><b>Manglet parametere.</b></span>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_EmailRemove(HttpListenerRequest request, int id)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: E-post", "?p=email");

                KgsaEmail email = new KgsaEmail(main);
                if (email.Remove(id.ToString()))
                {
                    Logg.Log("Web: Webbruker slettet E-post adresse med id " + id + ".", Color.DarkGoldenrod);
                    doc.Add("<span style='color:greem;'><b>E-post adressen med id " + id + " ble slette fra adresseboken.</b></span>");
                }
                else
                    doc.Add("<span style='color:orange;'><b>Kunne ikke slette e-posten fra adresseboken.</b></span>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_Email(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: E-post");

                KgsaEmail email = new KgsaEmail(main);
                if (email.emailDb != null)
                {
                    doc.Add("<div class='Table'><div class='Title'>Adresseboken</div>");
                    doc.Add("<div class='Heading'>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Navn</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>E-post adresse</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Gruppe</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Kveldstall</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Slett</p>");
                    doc.Add("</div>");
                    doc.Add("</div>");

                    for (int i = 0; i < email.emailDb.Rows.Count; i++)
                    {
                        doc.Add("<div class='Row'>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + email.emailDb.Rows[i]["Name"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + email.emailDb.Rows[i]["Address"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + email.emailDb.Rows[i]["Type"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + email.emailDb.Rows[i]["Quick"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p><a href='?p=email&action=remove&id=" + email.emailDb.Rows[i]["Id"] + "' onclick='return confirmRemove()'>Slett</a></p>");
                        doc.Add("</div>");
                        doc.Add("</div>");
                    }
                    doc.Add("</div>");
                }

                doc.Add("<br><br><form name=\"form1\" method=\"post\" action=\"?p=email&action=add\">");
                doc.Add("<div class='Table'><div class='Title'>Legg til</div>");
                doc.Add("<div class='Heading'>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Navn</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>E-post adresse</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Gruppe</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Kveldstall</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p>Slett</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='text' name='emailName' id='emailName'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='text' name='emailAddress' id='emailAddress'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><select name='emailGroup' id='emailGroup'>");
                doc.Add("<option value='Full'>Full</option><option value='AudioVideo'>AudioVideo</option><option value='Telecom'>Telecom</option><option value='Computer'>Computer</option><option value='Cross'>Cross</option>");
                doc.Add("</select></p></div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='checkbox' name='emailQuick' id='emailQuick'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell'>");
                doc.Add("<p><input type='submit' name='Lagre' id='submit' value='Submit'></p>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("</form>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings_Budget(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: Budsjett (WIP!)");

                if (main.budget != null)
                {
                    doc.Add("<div class='Table' style='width:690px;'><div class='Title'>Budsjett</div>");
                    doc.Add("<div class='Heading'>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Kategori</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Måned</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Omsetning</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Inntjening</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Margin</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Oppdatert</p>");
                    doc.Add("</div>");
                    doc.Add("<div class='Cell'>");
                    doc.Add("<p>Slett</p>");
                    doc.Add("</div>");
                    doc.Add("</div>");

                    var dtBudgets = main.budget.GetAllBudgets();

                    for (int i = 0; i < dtBudgets.Rows.Count; i++)
                    {
                        doc.Add("<div class='Row'>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + dtBudgets.Rows[i]["Kategori"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + Convert.ToDateTime(dtBudgets.Rows[i]["Date"]).ToString("MMMM yyyy") + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + dtBudgets.Rows[i]["Omsetning"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + dtBudgets.Rows[i]["Inntjening"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + dtBudgets.Rows[i]["Margin"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p>" + dtBudgets.Rows[i]["Updated"] + "</p>");
                        doc.Add("</div>");
                        doc.Add("<div class='Cell'>");
                        doc.Add("<p><a href='?p=budget&action=remove&id=" + dtBudgets.Rows[i]["Id"] + "' onclick='return confirmRemove()'>Slett</a></p>");
                        doc.Add("</div>");
                        doc.Add("</div>");
                    }
                    doc.Add("</div>");
                }
                else
                    doc.Add("<p>Budsjett ikke initialisert.</p>");

                doc.Add("<br><br><form name=\"form1\" method=\"post\" action=\"?p=budget&action=add\">");
                doc.Add("<div class='Table Settings'><div class='Title' style='color:white;'>Legg til nytt budsjett</div>");

                doc.Add("<div class='Heading'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:center;width:150px;'>");
                doc.Add("<p>Navn</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;text-align:center;'>");
                doc.Add("<p>Input</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;text-align:center;'>");
                doc.Add("<p>Beskrivelse</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Kategori
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Kategori:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none'>");
                doc.Add("<p><select name='budgetKat' id='budgetKat'>");
                doc.Add("<option value='MDA'>MDA</option><option value='AudioVideo'>AudioVideo</option><option value='SDA'>SDA</option><option value='Tele'>Tele</option><option value='Data'>Data</option><option value='Cross'>Cross</option><option value='Kasse'>Kasse</option><option value='Aftersales'>Aftersales</option>");
                doc.Add("</select></p></div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>Velg en kategori. Selgerkodene tilknyttet denne kategorien vil bli inkludert.</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Dato
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Velg måned:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none'>");
                doc.Add("<p><select name='budgetDate' id='budgetDate'>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(-1).ToShortDateString() + "'>" + DateTime.Now.AddMonths(-1).ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.ToShortDateString() + "' selected>" + DateTime.Now.ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(1).ToShortDateString() + "'>" + DateTime.Now.AddMonths(1).ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(2).ToShortDateString() + "'>" + DateTime.Now.AddMonths(2).ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(3).ToShortDateString() + "'>" + DateTime.Now.AddMonths(3).ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(4).ToShortDateString() + "'>" + DateTime.Now.AddMonths(4).ToString("MMMM yyyy") + "</option>");
                doc.Add("<option value='" + DateTime.Now.AddMonths(5).ToShortDateString() + "'>" + DateTime.Now.AddMonths(5).ToString("MMMM yyyy") + "</option>");
                doc.Add("</select></p></div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>Budsjett måneden.</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Antall åpningsdager
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Antall åpningsdager:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;'>");
                doc.Add("<p><input type='text' name='budgetDager' id='budgetDager' style='width:100px;'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>Antall åpningsdager i valgt måned. Må settes manuelt.</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Omsetning
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Omsetning:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;'>");
                doc.Add("<p><input type='text' name='budgetOmsetning' id='budgetOmsetning'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>&nbsp;</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Inntjening
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Inntjening:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;'>");
                doc.Add("<p><input type='text' name='budgetInntjening' id='budgetInntjening'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>&nbsp;</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Margin
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>Margin:</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;'>");
                doc.Add("<p><input type='text' name='budgetMargin' id='budgetMargin' style='width:100px;'></p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>&nbsp;</p>");
                doc.Add("</div>");
                doc.Add("</div>");

                // Lagre
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>&nbsp;</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;'>");
                doc.Add("<p>&nbsp;</p>");
                doc.Add("</div>");
                doc.Add("<div class='Cell' style='border:none;padding-bottom:5px;padding-top:5px;text-align:right;'>");
                doc.Add("<p><input type='submit' name='Lagre' id='submit' value='Legg til'></p>");
                doc.Add("</div>");
                doc.Add("</div>");

                doc.Add("</div>"); // table end
                doc.Add("<input type='hidden' name='returnPage' id='returnPage' value='settings?p=budget'>");
                doc.Add("</form>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private string WebPage_Settings(HttpListenerRequest request)
        {
            try
            {
                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger");

                doc.Add("Kontrollpanel: <a href='/settings?p=email'>Adresseboken</a> | <a href='/settings?p=vinn'>Vinnprodukter</a> | <a href='/settings?p=budget'>Budsjett (WIP!)</a><br><br>");

                // Setting table top
                doc.Add("<form id='Settings' name='Settings' method='post' action='?p=save'>");
                doc.Add("<div class='Table Settings'><div class='Title' style='color:white;'>Generelt</div>");

                doc.AddRange(CreateRowSettingFor("rankingCompareLastyear", "Sammenlign i fjor:", "", "Av", "På", "Utvidet"));
                doc.AddRange(CreateRowSettingFor("rankingCompareLastmonth", "Sammenlign sist måned:", "", "Av", "På", "Utvidet"));
                doc.AddRange(CreateRowSettingFor("showInfo", "Vis infotekst:", "Viser introduksjons tekst øverst på hvert enkelt ranking side."));

                // Settings table bottom
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='border:none;'>&nbsp;</div>");
                doc.Add("<div class='Cell' style='border:none;'>&nbsp;</div>");
                doc.Add("<div class='Cell' style='border:none;padding-bottom:5px;padding-top:5px;text-align:right;'><input type='submit' name='Lagre' id='submit' value='Lagre'></div>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("<input type='hidden' name='returnPage' id='returnPage' value='settings'>");
                doc.Add("</form>");
                
                // Actions table top
                doc.Add("<div class='Table'>");
                doc.Add("<div class='Heading'>");
                doc.Add("<div class='Cell'>Handlinger:");
                doc.Add("</div>");
                doc.Add("</div>");

                // Actions table body
                doc.Add("<div class='Row'>");
                doc.Add("<div class='Cell' style='padding-left:10px;'>");
                doc.Add("<p><a href=\"?p=process&action=update\" onclick=\"return confirmAction()\">Oppdater alle sider</a><br>");
                doc.Add("Tips: Bruk hvis noen sider ikke er oppdatert selv om data er importert fra Elguide.<br><br>");
                doc.Add("<a href=\"?p=process&action=importranking\" onclick=\"return confirmAction()\">Start makro-importering av transaksjoner nå</a><br>");
                doc.Add("Tips: Importer fra Elguide hvis ranking ikke er oppdatert.<br><br>");
                doc.Add("<a href=\"?p=process&action=importstore\" onclick=\"return confirmAction()\">Start makro-importering av lager nå</a><br>");
                doc.Add("Tips: Importer fra Elguide hvis lager ikke er oppdatert.<br><br>");
                doc.Add("<a href=\"?p=process&action=importservice\" onclick=\"return confirmAction()\">Start makro-importering av servicer nå</a><br>");
                doc.Add("Tips: Importer fra Elguide hvis servicer ikke er oppdatert.<br><br>");
                doc.Add("<a href=\"?p=process&action=autoranking\" onclick=\"return confirmAction()\">Start automatisk oppdatering og utsending av eposter</a><br>");
                doc.Add("Tips: Hvis dagens tjeneste-ranking ikke har blitt utsendt kan et nytt forsøk startes her.<br><br>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("</div>");
                doc.Add("<p><span style='color:orange;'>Advarsel: Når en handling er startet er det ikke mulig å stoppe prosessen fra dette kontrollpanelet.</span></p>");

                WebMenuEnd(doc);
                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return "";
        }

        private string WebPage_Settings_Save(HttpListenerRequest request)
        {
            try
            {
                Dictionary<string, string> postParams = ParseParameters(request);

                string returnPage = postParams.ContainsKey("returnPage") == true ? postParams["returnPage"] : "";

                List<string> doc = new List<string>() { };
                WebMenuStart(doc, "Innstillinger: Lagre", returnPage);

                bool settingsChanged = false;
                foreach(var key in postParams.Keys)
                {
                    foreach (var prop in main.appConfig.GetType().GetProperties())
                    {
                        if (prop.Name == key)
                        {
                            var oldval = prop.GetValue(main.appConfig, null);
                            if (protectedSettings.Contains(prop.Name))
                            {
                                Logg.Log("Web: Webbruker forsøkte å endre beskyttede variabler! (" + prop.Name + ")", Color.Red);
                                goto breakloop;
                            }

                            if (typeof(string).IsAssignableFrom(prop.PropertyType))
                            {
                                string value = System.Web.HttpUtility.UrlDecode(postParams[key]);
                                if (value != (string)oldval)
                                {
                                    prop.SetValue(main.appConfig, Convert.ChangeType(value, prop.PropertyType), null);
                                    Logg.Log("Web: Webbruker satt ny verdi på appConfig." + prop.Name + " = " + value, Color.DarkGoldenrod);
                                    settingsChanged = true;
                                }
                            }
                            else if (typeof(bool).IsAssignableFrom(prop.PropertyType))
                            {
                                bool value;
                                if (bool.TryParse(postParams[key], out value))
                                {
                                    if (value != (bool)oldval)
                                    {
                                        prop.SetValue(main.appConfig, Convert.ChangeType(value, prop.PropertyType), null);
                                        Logg.Log("Web: Webbruker satt ny verdi på appConfig." + prop.Name + " = " + value, Color.DarkGoldenrod);
                                        settingsChanged = true;
                                    }
                                }
                                else
                                    Logg.Debug("Web: En innstilling ble ikke lagret riktig. (" + prop.Name + ")");
                            }
                            else if (typeof(decimal).IsAssignableFrom(prop.PropertyType))
                            {
                                decimal value;
                                string valueStr = System.Web.HttpUtility.UrlDecode(postParams[key]);
                                if (decimal.TryParse(valueStr, out value))
                                {
                                    if (value != (decimal)oldval)
                                    {
                                        prop.SetValue(main.appConfig, Convert.ChangeType(value, prop.PropertyType), null);
                                        Logg.Log("Web: Webbruker satt ny verdi på appConfig." + prop.Name + " = " + value, Color.DarkGoldenrod);
                                        settingsChanged = true;
                                    }
                                }
                                else
                                    Logg.Debug("Web: En innstilling ble ikke lagret riktig. (" + prop.Name + ")");
                            }
                            else if (typeof(int).IsAssignableFrom(prop.PropertyType))
                            {
                                int value;
                                if (int.TryParse(postParams[key], out value))
                                {
                                    if (value != (int)oldval)
                                    {
                                        prop.SetValue(main.appConfig, Convert.ChangeType(value, prop.PropertyType), null);
                                        Logg.Log("Web: Webbruker satt ny verdi på appConfig." + prop.Name + " = " + value, Color.DarkGoldenrod);
                                        settingsChanged = true;
                                    }
                                }
                                else
                                    Logg.Debug("Web: En innstilling ble ikke lagret riktig. (" + prop.Name + ")");
                            }
                            else
                            {
                                Logg.Debug("Web: En innstilling ble ikke lagret riktig. Type ikke godkjent. (" + prop.Name + ")");
                            }
                        }
                    }
                }

                if (settingsChanged)
                {
                    doc.Add("<span style='color:green;'><b>Innstillinger endret.</b></span>");
                    main.SaveSettings();
                }
                else
                {
                    Logg.Log("Web: Webbruker forsøkte å lagre innstillinger, men ingen innstillinger ble endret.", Color.DarkGoldenrod);
                    doc.Add("<span style='color:orange;'><b>Ingen innstillinger be endret.</b></span>");
                }

                breakloop:

                WebMenuEnd(doc);

                return string.Join(Environment.NewLine, doc.ToArray());
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return "";
        }

        private string[] protectedSettings = new string[1] { "webserverPassword" };

        private Dictionary<string, string> ParseParameters(HttpListenerRequest request)
        {
            try
            {
                string str = GetRequestPostData(request);
                Dictionary<string, string> postParams = new Dictionary<string, string>();
                string[] rawParams = str.Split('&');
                foreach (string param in rawParams)
                {
                    string[] kvPair = param.Split('=');
                    string key = kvPair[0];
                    string value = kvPair[1];
                    postParams.Add(key, value);
                }

                return postParams;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private List<string> CreateRowSettingFor(string field, string description = "", string info = "",
    string option0 = "", string option1 = "", string option2 = "", string option3 = "", string option4 = "", string option5 = "")
        {
            List<string> doc = new List<string>() { };
            doc.Add("<div class='Row'>");

            if (!String.IsNullOrEmpty(description))
            {
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>" + description + "</p>");
                doc.Add("</div>");
            }
            else
            {
                doc.Add("<div class='Cell' style='border:none;padding-top:3px;padding-bottom:3px;padding-left:3px;padding-right:10px;text-align:right;width:150px;'>");
                doc.Add("<p>" + field + "</p>");
                doc.Add("</div>");
            }

            foreach (var prop in main.appConfig.GetType().GetProperties())
            {
                if (prop.Name == field)
                {
                    if (typeof(string).IsAssignableFrom(prop.PropertyType))
                    {
                        string val = (string)prop.GetValue(main.appConfig, null);
                        doc.Add("<div class='Cell' style='border:none'>");
                        doc.Add("<p><input type='text' name='" + prop.Name + "' id='" + prop.Name + "' value='" + val + "'></p>");
                        doc.Add("</div>");
                    }
                    if (typeof(bool).IsAssignableFrom(prop.PropertyType))
                    {
                        bool val = (bool)prop.GetValue(main.appConfig, null);
                        doc.Add("<div class='Cell' style='border:none;'>");
                        doc.Add("<p><select name='" + prop.Name + "' id='" + prop.Name + "'>");
                        if (val)
                        {
                            doc.Add("<option value='true' selected>Sant</option>");
                            doc.Add("<option value='false'>Ikke sant</option>");
                        }
                        else
                        {
                            doc.Add("<option value='true'>Sant</option>");
                            doc.Add("<option value='false' selected>Ikke sant</option>");
                        }
                        doc.Add("</select></p></div>");
                    }
                    if (typeof(int).IsAssignableFrom(prop.PropertyType))
                    {
                        int val = (int)prop.GetValue(main.appConfig, null);
                        if (!String.IsNullOrEmpty(option0))
                        {
                            doc.Add("<div class='Cell' style='border:none;'>");
                            doc.Add("<p><select name='" + prop.Name + "' id='" + prop.Name + "'>");

                            if (val == 0)
                                doc.Add("<option value='0' selected>" + option0 + "</option>");
                            else
                                doc.Add("<option value='0'>" + option0 + "</option>");
                            if (!String.IsNullOrEmpty(option1))
                            {
                                if (val == 1)
                                    doc.Add("<option value='1' selected>" + option1 + "</option>");
                                else
                                    doc.Add("<option value='1'>" + option1 + "</option>");
                            }
                            if (!String.IsNullOrEmpty(option2))
                            {
                                if (val == 2)
                                    doc.Add("<option value='2' selected>" + option2 + "</option>");
                                else
                                    doc.Add("<option value='2'>" + option2 + "</option>");
                            }
                            if (!String.IsNullOrEmpty(option3))
                            {
                                if (val == 3)
                                    doc.Add("<option value='3' selected>" + option3 + "</option>");
                                else
                                    doc.Add("<option value='3'>" + option3 + "</option>");
                            }
                            if (!String.IsNullOrEmpty(option4))
                            {
                                if (val == 4)
                                    doc.Add("<option value='4' selected>" + option4 + "</option>");
                                else
                                    doc.Add("<option value='4'>" + option4 + "</option>");
                            }
                            if (!String.IsNullOrEmpty(option5))
                            {
                                if (val == 5)
                                    doc.Add("<option value='5' selected>" + option5 + "</option>");
                                else
                                    doc.Add("<option value='5'>" + option5 + "</option>");
                            }

                            doc.Add("</select></p></div>");
                        }
                        else
                        {
                            doc.Add("<div class='Cell' style='border:none'>");
                            doc.Add("<p><input type='text' name='" + prop.Name + "' id='" + prop.Name + "' value='" + val + "'></p>");
                            doc.Add("</div>");
                        }
                    }
                    if (typeof(decimal).IsAssignableFrom(prop.PropertyType))
                    {
                        decimal val = (decimal)prop.GetValue(main.appConfig, null);
                        doc.Add("<div class='Cell' style='border:none'>");
                        doc.Add("<p><input type='text' name='" + prop.Name + "' id='" + prop.Name + "' value='" + val + "'></p>");
                        doc.Add("</div>");
                    }
                }
            }

            if (!String.IsNullOrEmpty(info))
            {
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>");
                doc.Add("<p>" + info + "</p>");
                doc.Add("</div>");
            }
            else
                doc.Add("<div class='Cell' style='border:none;padding-left:5px;padding-right:3px;'>&nbsp;</div>");

            doc.Add("</div>");
            return doc;
        }

        public static string GetRequestPostData(HttpListenerRequest request)
        {
            if (!request.HasEntityBody)
            {
                return null;
            }
            using (System.IO.Stream body = request.InputStream) // here we have data
            {
                using (System.IO.StreamReader reader = new System.IO.StreamReader(body, request.ContentEncoding))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static string WebPage_FileNotFound()
        {
            List<string> doc = new List<string>() { };
            GetHtmlStartWeb(doc, "KGSA - 404");

            doc.Add("<div class='container'><div class='header'>KGSA " + FormMain.version + " (Build: " + FormMain.RetrieveLinkerTimestamp().ToShortDateString() + ")</div>");
            doc.Add("<div class='container-inside'>");
            doc.Add("<div class='container-content'>");
            doc.Add("<span class='breadcrum'><a href='/'>Meny</a>: 404</span>");
            doc.Add("<div style=\"clear: both;\"></div>");
            doc.Add("<div class=\"container-menu\">");
            doc.Add("Vi fant ikke siden du ønsket.<br>Gå tilbake til <a href='/'>forsiden?</a>");
            doc.Add("</div>");
            doc.Add("</div>");
            doc.Add("</div>");
            doc.Add("</div>");

            doc.Add(Resources.htmlEnd);

            return string.Join(Environment.NewLine, doc.ToArray());
        }

        /// <summary>
        /// Starten av web menyen
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="title">Tittel på siden som vises også i breadcrum menyen</param>
        /// <param name="redirect">Omdiriger til denne siden etter 5 sekund</param>
        /// <param name="autorefresh">Oppfrisk siden hvert 5. sekund</param>
        private void WebMenuStart(List<string> doc, string title, string redirect = "", bool autorefresh = false)
        {
            GetHtmlStartWeb(doc, "KGSA - " + title, redirect, autorefresh);
            doc.Add("<div class='container'><div class='header'>KGSA " + FormMain.version + " (Build: " + FormMain.RetrieveLinkerTimestamp().ToShortDateString() + ")</div>");
            doc.Add("<div class='container-inside'>");
            doc.Add("<div class='container-content'>");
            doc.Add("<span class='breadcrum'><a href='/'>Meny</a>: " + title + "</span>");
            doc.Add("<div style=\"clear: both;\"></div>");
            doc.Add("<div class=\"container-menu\">");
        }

        private void WebMenuEnd(List<string> doc)
        {
            doc.Add("</div>");
            doc.Add("</div>");
            doc.Add("</div>");
            doc.Add("</div>");
            doc.Add(Resources.htmlEnd.Replace("<!-- jsWeb here -->", Resources.jsWeb));
        }

        private static void GetHtmlStartWeb(List<string> doc, string title = "KGSA Web", string redirect = "", bool autorefresh = false)
        {
            doc.Add("<html>");
            doc.Add("<head>");
            doc.Add("<meta charset=\"UTF-8\">");
            if (!String.IsNullOrEmpty(redirect))
                doc.Add("<meta http-equiv=\"refresh\" content=\"3; url=" + redirect + "\" />");
            if (autorefresh)
                doc.Add("<meta http-equiv=\"refresh\" content=\"8\" >");

            doc.Add("<title>" + title + "</title>");

            doc.Add("<style id=\"stylesheet\">");
            doc.Add("body {");
            doc.Add("    font-weight:400;");
            doc.Add("    font-family:Calibri, sans-serif;");
            doc.Add("    font-style:normal;");
            doc.Add("    text-decoration:none;");
            doc.Add("    display: block;");
            doc.Add("    padding: 0px;");
            doc.Add("    margin: 15px;");
            doc.Add("}");
            doc.Add("</style>");
            doc.Add(Resources.webstyle);
            doc.Add("</head>");
            doc.Add("<body>");
        }
    }
}
