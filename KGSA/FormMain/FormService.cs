using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    partial class FormMain
    {
        public void ReloadService(bool forced = false)
        {
            if (forced)
            {
                ClearHash("ServiceOversikt");
                ClearHash("ServiceList");
            }

            if (service.dbServiceDatoFra == service.dbServiceDatoTil)
            {
                labelServiceHead.ForeColor = SystemColors.ControlText;
                labelServiceDate.ForeColor = SystemColors.ControlText;
                labelServiceHead.Text = "(tom)";
                labelServiceDate.Text = "";
                webService.Navigate(htmlImportService);

                ShowHideGui_EmptyService(false);
            }
            else
            {
                labelServiceHead.Text = service.dbServiceDatoTil.ToString("dddd", norway);
                labelServiceDate.Text = service.dbServiceDatoTil.ToString("d. MMM", norway);

                ShowHideGui_EmptyService(true);

                if ((DateTime.Now - service.dbServiceDatoTil).Days >= 3)
                {
                    labelServiceHead.ForeColor = Color.Red;
                    labelServiceDate.ForeColor = Color.Red;
                }
                else
                {
                    labelServiceHead.ForeColor = SystemColors.ControlText;
                    labelServiceDate.ForeColor = SystemColors.ControlText;
                }

                if (!autoMode)
                    UpdateRank();
                string page = currentServicePage();
                if (page == "ServiceOversikt")
                    RunServiceOversikt();
                if (page == "ServiceList")
                    RunServiceList();
                else
                    RunServiceOversikt();

                moveDateService(0);
            }
        }

        public void RetrieveDbService()
        {
            service.Load(this);

            if (service.dbServiceDatoFra.Date == service.dbServiceDatoTil.Date)
            {
                ChangeServiceDateTimePicker(DateTime.Now, rangeMin, rangeMax);
            }
            else
            {
                ChangeServiceDateTimePicker(service.dbServiceDatoTil, service.dbServiceDatoFra, service.dbServiceDatoTil);
                Logg.Log("Service databasen har servicer mellom " + service.dbServiceDatoFra.ToString("d. MMMM yyyy", norway) + " og " + service.dbServiceDatoTil.ToString("d. MMMM yyyy", norway) + " for din avdeling.", Color.Black, true);
            }

        }

        private void RunServiceImport(string str = "")
        {
            if (!IsBusy(true))
            {
                processing.SetVisible = true;
                processing.SetText = "Importerer servicer..";
                Logg.Log("Importerer servicer..");
                processing.SetValue = 25;
                processing.SetBackgroundWorker = bwImportService;
                bwImportService.RunWorkerAsync(str);
            }
        }

        private void RunAutoServiceImport(string str = "")
        {
            if (!IsBusy(true))
            {
                processing.SetVisible = true;
                processing.SetText = "Importerer servicer med makro..";
                Logg.Log("Importerer servicer med makro..");
                processing.SetValue = 25;
                processing.SetBackgroundWorker = bwAutoImportService;
                bwAutoImportService.RunWorkerAsync(str);
            }
        }

        private void delayedAutoServiceImport()
        {
            try
            {
                processing.SetVisible = true;
                processing.SetProgressStyle = ProgressBarStyle.Continuous;
                processing.SetBackgroundWorker = bwAutoImportService;
                for (int b = 0; b < 100; b++)
                {
                    processing.SetText = "Starter automatisk service import om " + (((b / 10) * -1) + 10) + " sekunder..";
                    processing.SetValue = b;
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);

                    if (processing.userPushedCancelButton)
                    {
                        Logg.Log("Brukeren avbrøt handlingen.");
                        return;
                    }
                }

                RestoreWindow();

                RunAutoServiceImport();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwImportService_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string str = (string)e.Argument;
            bool complete = false;

            if (!String.IsNullOrEmpty(str))
                complete = service.Import(str, processing, bwImportService);
            else
                complete = service.Import(appConfig.csvElguideExportFolder + @"\iserv.csv", processing, bwImportService);

            e.Result = complete;
        }

        private void bwProgressCustom_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ReportProgressCustom(Convert.ToDecimal(e.ProgressPercentage), (StatusProgress)e.UserState);
        }

        private void bwImportService_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            try
            {
                bool complete = (bool)e.Result;
                if (complete)
                {
                    Logg.Log("Service Import: Service import fullført uten feil.");
                    database.ClearCacheTables();
                    RetrieveDbService();
                    ReloadService();
                    UpdateUi();
                    UpdateService();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            processing.HideDelayed();
            this.Activate();


        }

        public void UpdateTimerService()
        {
            try
            {
                if (appConfig.AutoService)
                {
                    TimeSpan check = timerNextRunService.Subtract(DateTime.Now);
                    SetStatusInfo("timer", "service", "", timerNextRunService);

                    if (check.TotalMinutes < 1 && check.TotalMinutes >= 0)
                    {
                        if (!IsBusy(true))
                            delayedAutoServiceImport();
                        return;
                    }

                    if (check.TotalMinutes < 0)
                    {
                        DateTime s = timerNextRunService;
                        var ts = TimeSpan.FromMinutes(appConfig.serviceAutoImportMinutter);
                        var tFra = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, appConfig.serviceAutoImportFraIndex, 0, 0);
                        var tTil = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, appConfig.serviceAutoImportTilIndex, 0, 0);

                        if (DateTime.Now > tTil)
                        {
                            tFra = tFra.AddDays(1);
                            tTil = tTil.AddDays(1);
                        }

                        int limit = 50;
                        do
                        {
                            limit--;
                            timerNextRunService = timerNextRunService.AddMinutes(appConfig.serviceAutoImportMinutter);
                        }
                        while (!(timerNextRunService < tTil && timerNextRunService > tFra) && limit > 0);

                        if (timerNextRunService.DayOfWeek == DayOfWeek.Sunday && appConfig.ignoreSunday)
                            timerNextRunService = timerNextRunService.AddDays(1);

                        Logg.Debug("Neste service import: " + timerNextRunService);
                        return;
                    }

                    if (check.TotalMinutes < 15 && check.TotalMinutes > 1)
                    {
                        if (Math.Round(check.TotalMinutes) == 1)
                            Logg.Log("Starter automatisk innhenting av servicer om 1 minutt.");
                        else
                            Logg.Log("Starter automatisk innhenting av servicer om " + Math.Round(check.TotalMinutes - 1) + " minutter.");
                        return;
                    }
                }
                else
                {
                    timerAutoService.Stop();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwAutoImportService_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ProgressStart();
                autoMode = true;
                var macroAttempt = 1;
                var macroMaxAttempts = 4;

                retrymacro:

                Logg.Log("AutoService: Kjører Service makro.. [" + macroAttempt + "]");
                var macroForm = (FormMacro)StartMacro(DateTime.Now, macroProgramService, bwAutoImportService, macroAttempt);
                if (macroForm.errorCode != 0)
                {
                    // Feil oppstod under kjøring av macro
                    macroAttempt++;
                    Logg.Log("Auto: Feil oppstod under kjøring av makro. Feilbeskjed: " + macroForm.errorMessage + " Kode: " + macroForm.errorCode, Color.Red);
                    for (int i = 0; i < 60; i++)
                    {
                        Application.DoEvents();
                        System.Threading.Thread.Sleep(50);
                    }
                    if (macroForm.errorCode == 6)
                        return;
                    if (macroAttempt < macroMaxAttempts && macroForm.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                        goto retrymacro;

                    return;
                }

                if (macroForm.errorCode == 0)
                {
                    service.Load(this);

                    if (service.Import(appConfig.csvElguideExportFolder + "iserv.csv", processing, bwAutoImportService))
                    {
                        Logg.Log("AutoService: Service import fullført uten feil.");
                        MakeServiceOversikt(true, bwAutoImportService);
                        Logg.Log("AutoService: Service rapport laget.");
                    }
                }

                Logg.Log("AutoService: Ferdig.");
                e.Result = macroForm.errorCode;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwAutoImportService_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;

            if (e.Cancelled)
            {
                Logg.Log("Auto: Jobb kansellert.");
            }
            else if (e.Error != null)
            {
                Logg.Log("Auto: Feil oppstod under kjøring av makro. Se logg for detaljer. (" + e.Error.Message + ")");
            }
            else
            {
                int returnCode = 2;
                if (e.Result != null)
                    returnCode = (int)e.Result;

                if (returnCode == 0)
                {
                    RetrieveDbService();
                    ReloadService();
                    UpdateUi();
                    Logg.Log("Auto: Makro fullført.");
                }
                else if (returnCode == 2)
                {
                    Logg.Log("Auto: Jobb avbrutt av bruker.");
                }
            }

            processing.HideDelayed();
            this.Activate();
        }

        private void UpdateService(string katArg = "")
        {
            if (service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
            {
                string page = currentServicePage();
                if (String.IsNullOrEmpty(katArg) && !String.IsNullOrEmpty(page))
                    ClearHash(page);
                
                if (page == "ServiceOversikt")
                    RunServiceOversikt();
                else if (page == "ServiceList" || page == "ServiceDetails")
                    RunServiceList();
                else
                    RunServiceOversikt();
            }
            else
            {
                webService.Navigate(htmlImportService);
            }
        }

        private string currentServicePage()
        {
            if (webService.Url != null)
            {
                string str = webService.Url.OriginalString;
                if (str.Contains("serviceOversikt.html"))
                    return "ServiceOversikt";
                else if (str.Contains("serviceList.html"))
                    return "ServiceList";
                else if (str.Contains("serviceDetails.html"))
                    return "ServiceDetails";
                return "ServiceOversikt";
            }
            return "";
        }

        private void moveDateService(int m = 0)
        {
            if (service != null)
            {
                if (service.dbServiceDatoFra.Date == service.dbServiceDatoTil.Date)
                    return;

                var d = pickerServiceDato.Value;
                if (m == 1) // gå tilbake en måned
                {
                    if (service.dbServiceDatoFra.Date <= d.AddMonths(-1))
                        pickerServiceDato.Value = d.AddMonths(-1);
                    else
                        pickerServiceDato.Value = service.dbServiceDatoFra;
                }
                if (m == 2) // gå tilbake en dag
                {
                    if (appConfig.ignoreSunday)
                    {
                        if (service.dbServiceDatoFra.Date <= d.AddDays(-1) && d.AddDays(-1).DayOfWeek != DayOfWeek.Sunday)
                            pickerServiceDato.Value = d.AddDays(-1);
                        if (service.dbServiceDatoFra.Date <= d.AddDays(-2) && d.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
                            pickerServiceDato.Value = d.AddDays(-2);
                    }
                    else
                    {
                        if (service.dbServiceDatoFra.Date <= d.AddDays(-1))
                            pickerServiceDato.Value = d.AddDays(-1);
                    }
                }
                if (m == 3) // gå fram en dag
                {
                    if (appConfig.ignoreSunday)
                    {
                        if (service.dbServiceDatoTil.Date >= d.AddDays(1) && d.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                            pickerServiceDato.Value = d.AddDays(1);
                        if (service.dbServiceDatoTil.Date >= d.AddDays(2) && d.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            pickerServiceDato.Value = d.AddDays(2);
                    }
                    else
                    {
                        if (service.dbServiceDatoTil.Date >= d.AddDays(1))
                            pickerServiceDato.Value = d.AddDays(1);
                    }
                }
                if (m == 4) // gå fram en måned
                {
                    if (service.dbServiceDatoTil.Date >= d.AddMonths(1))
                        pickerServiceDato.Value = d.AddMonths(1);
                    else
                        pickerServiceDato.Value = service.dbServiceDatoTil;
                }
                d = pickerServiceDato.Value;
                if (d.Date >= service.dbServiceDatoTil.Date)
                {
                    buttonServF.Enabled = false; // fremover knapp
                    buttonServFF.Enabled = false; // fremover knapp
                }
                else
                {
                    buttonServF.Enabled = true; // fremover knapp
                    buttonServFF.Enabled = true; // fremover knapp
                }
                if (d.Date <= appConfig.dbFrom.Date)
                {
                    buttonServBF.Enabled = false; // bakover knapp
                    buttonServB.Enabled = false; // bakover knapp
                }
                else
                {
                    buttonServBF.Enabled = true; // bakover knapp
                    buttonServB.Enabled = true; // bakover knapp
                }

                UpdateService();
            }
        }

        private void velgServiceCSV()
        {
            // Browse etter iserv.csv
            try
            {
                var fdlg = new OpenFileDialog();
                fdlg.Title = "Velg Servce CVS-fil eksportert fra Elguide";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "All files (*.*)|*.*|CVS filer (*.csv)|*.csv";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                fdlg.Multiselect = false;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    RunServiceImport(fdlg.FileName);
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void MakeServiceDetails(int serviceID)
        {
            Logg.Debug("Søker service logg..");
            Logg.Status("Søker service logg..");
            try
            {
                var doc = new List<string>();

                GetHtmlStart(doc, true);

                doc.Add("<span class='Title'>Service detaljer (" + avdeling.Get(appConfig.Avdeling) + ")</span><span class='Generated'>Side generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");
                doc.Add("<br><table style='width:100%'><tr><td>");

                service.GenerateServiceDetails(doc, serviceID);
                doc.Add("</td></tr></table>");
                doc.Add(Resources.htmlEnd);

                File.WriteAllLines(htmlServiceDetails, doc.ToArray(), Encoding.Unicode);
                webService.Navigate(htmlServiceDetails);
            }
            catch (Exception ex)
            {
                webService.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av service detaljer.", ex);
                errorMsg.ShowDialog(this);
            }
            Logg.Status("Klar.");
        }

        public void RunServiceList(string statusFilter = "", string loggFilter = "")
        {
            if (service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date && !bwServiceReport.IsBusy && !bwService.IsBusy)
            {
                object param1 = statusFilter;
                object param2 = loggFilter;
                object[] parameters = new object[] { param1, param2 };
                bwServiceReport.RunWorkerAsync(parameters);
                HighlightServiceButton("ServiceList");
            }
            else
                webService.Navigate(htmlImportService);
        }

        private void bwServiceReport_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            object[] parameters = e.Argument as object[];
            string statusFilter = parameters[0].ToString();
            string loggFilter = parameters[1].ToString();
            string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
            MakeServiceList(statusFilter, loggFilter, false, bwService);
            appConfig.strServiceList = newHash;
        }

        private void MakeServiceList(string statusFilter = "", string loggFilter = "", bool bg = false, BackgroundWorker bw = null)
        {
            Logg.Log("Lager [ServiceList] ranking..");
            bool abort = HarSisteVersjonService("ServiceList", appConfig.strServiceList);
            try
            {
                var doc = new List<string>();

                GetHtmlStart(doc, true);

                doc.Add("<span class='Title'>Service oversikt  (" + avdeling.Get(appConfig.Avdeling) + ")</span><span class='Generated'>Side generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                var statusObject = new ServiceStatus();
                var fieldValues = statusObject.GetType()
                     .GetFields()
                     .Select(field => field.GetValue(statusObject))
                     .ToList();

                string statusNavn = statusFilter;
                if (statusFilter.Length == 0)
                    statusNavn = "Alle";
                else
                    statusFilter = statusNavn;
                doc.Add("<br><span class='Subtitle'>Status: <b>" + statusNavn + "</b></span>");
                doc.Add("<br><form action=\"#\" name='ex' style=\"display: inline;\"><select name=\"status\" onchange=\"window.open(document.ex.status.options[document.ex.status.selectedIndex].value,'_top')\">");
                doc.Add("<option value=''>Velg status..</option>");
                doc.Add("<option value='#list_alle'>Alle aktive</option>");
                foreach (string status in fieldValues)
                    if (!status.Contains("Ferdig"))
                        doc.Add("<option value='#list_" + status.Replace(" ", "%20") + "'>" + status + "</option>");
                doc.Add("</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class='Subtext' style=\"display: inline;\">Logg filter:</span>&nbsp;&nbsp;");

                doc.Add("<input type=\"text\" name=\"state\" id=\"state\" value=\"\" size=\"10\" />");
                doc.Add("<select name=\"s\" size=\"1\" \">");
                doc.Add("<option value=\" \" selected=\"selected\"></option>");
                doc.Add("<option value=\"#search_Venter%20svar\">Venter svar</option>");
                doc.Add("</select>&nbsp;&nbsp;<input id=\"clickMe\" type=\"submit\" value=\"Søk\"></form>");

                doc.Add("<br><br><table style='width:100%'><tr><td>");
                doc.Add("<span class='Subtitle'>Aktive servicer (oppdatert " + service.dbServiceDatoTil.ToString("dddd d. MMMM yyyy", norway) + ")</span>");
                service.GenerateServiceList(statusFilter, loggFilter, doc, bw);
                doc.Add("</td></tr></table>");
                doc.Add(Resources.htmlEnd);

                File.WriteAllLines(htmlServiceList, doc.ToArray(), Encoding.Unicode);
                if (!bg)
                    webService.Navigate(htmlServiceList);

            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                if (!bg)
                {
                    webService.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av service oversikt.", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void bwServiceReport_Completed(object sender, AsyncCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Logg.Status("Klar.");
        }

        private void RunServiceOversikt()
        {
            if (service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date && !bwServiceReport.IsBusy && !bwService.IsBusy)
            {
                HighlightServiceButton("ServiceOversikt");
                bwService.RunWorkerAsync("ServiceOversikt");
            }
            else
                webService.Navigate(htmlImportService);
        }

        private void bwService_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
            MakeServiceOversikt(false, bwService);
            appConfig.strServiceOversikt = newHash;
        }

        private void MakeServiceOversikt(bool bg = false, BackgroundWorker bw = null)
        {
            string katArg = "ServiceOversikt";
            bool abort = HarSisteVersjonService("ServiceOversikt", appConfig.strServiceOversikt);
            try
            {
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webService.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerServiceDato.Value;

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Service oversikt (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    doc.Add("<h3>Status servicer</h3>");
                    service.GenerateServiceOversikt(doc, bw, dtPick);

                    doc.Add("<h3>Status servicer</h3>");
                    service.GenerateServiceEssentials(doc);

                    if (appConfig.favVis)
                    {
                        doc.Add("<h3>Favoritt avdelinger</h3>");
                        service.GenerateFavorittAvdelinger(doc, bw, dtPick);
                    }

                    if (appConfig.serviceShowHistory)
                    {
                        doc.Add("<h3>Service historikk</h3>");
                        service.GenerateServiceHistory(doc);
                    }

                    if (appConfig.serviceShowHistoryGraph)
                    {
                        doc.Add("<h3>Service historikk</h3>");
                        service.GenerateServiceHistoryGraph(doc, bw, dtPick);
                    }

                    if (appConfig.serviceFerdigServiceStats)
                    {
                        doc.Add("<h3>Fullførte servicer siste " + appConfig.serviceHistoryDays + " dager pr. selgerkode</h3>");
                        service.GenerateServiceFerdigStats(doc, bw, dtPick);
                    }

                    doc.Add(Resources.htmlEnd);

                    File.WriteAllLines(htmlServiceOversikt, doc.ToArray(), Encoding.Unicode);
                    if (!bg)
                        webService.Navigate(htmlServiceOversikt);
                }
                else if (!bg)
                    webService.Navigate(htmlServiceOversikt);
            }
            catch (Exception ex)
            {
                if (!bg)
                {
                    webService.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
                else
                    Logg.Unhandled(ex);
            }
        }

        private void bwServiceGraph_Completed(object sender, AsyncCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Logg.Status("Klar");
        }
    }


    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public class ScriptInterface
    {
        public void callBehandlet(string serviceId, string text)
        {
            try
            {
                Logg.Debug("Merk behandlet service id: " + serviceId + " | text: " + text);
                int varBit = 0;
                if (text == "Nei")
                    varBit = 1;

                string sql = "UPDATE tblService SET [FerdigBehandlet] = " + varBit + " WHERE [ServiceID] = " + serviceId;

                var con = new SqlCeConnection(FormMain.SqlConStr);
                var command = new SqlCeCommand(sql, con);
                con.Open();
                int var = command.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Unntak ved callBehandlet()", ex);
                errorMsg.ShowDialog();
            }
        }

        public void callBrowseServicePage(string s, string t = "")
        {
            UpdateServicePage.GetServicePage(s, t);
        }

        public void callBrowseList(string status, string filter)
        {
            UpdateServicePage.GetServiceList(status, filter);
        }
    }
}
