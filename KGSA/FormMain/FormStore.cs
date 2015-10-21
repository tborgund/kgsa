using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace KGSA
{
    partial class FormMain
    {
        private void RunAutoObsoleteImport()
        {
            if (!IsBusy(true))
            {
                processing.SetVisible = true;
                processing.SetText = "Importerer lager med makro..";
                Logg.Log("Importerer lager med makro..");
                processing.SetValue = 25;
                processing.SetBackgroundWorker = bwAutoStore;
                bwAutoStore.RunWorkerAsync();
            }
        }

        public bool EmptyStoreDatabase()
        {
            try
            {
                if (appConfig.dbStoreFrom.Date.Equals(DateTime.Now.Date))
                    return true;
                if (appConfig.dbStoreFrom == DateTime.MinValue || appConfig.dbStoreFrom == rangeMin)
                    return true;
                if (appConfig.dbStoreFrom > appConfig.dbStoreTo)
                    return true;
                else
                    return false; // OK det ser ut som lager databasen er IKKE tom.
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return true;
            }
        }

        private void ImportWobsoleteCsvZip()
        {
            try
            {
                string extracted = obsolete.Decompress(appConfig.csvElguideExportFolder + @"\wobsolete.zip");
                if (!String.IsNullOrEmpty(extracted))
                    RunObsoleteImport(extracted);
                else
                {
                    Logg.Log("Mislykkes forsøk ved utpakking av (" + appConfig.csvElguideExportFolder + @"\wobsolete.csv" + ")");
                    MessageBox.Show("Obs! Utpakking av arkiv mislykkes eller wobsolete.zip ble ikke funnet. Se logg for detaljer.\nPrøv igjen eller pakk ut manuelt før importering.", "KGSA - Info", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (FileNotFoundException ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Fant ikke fil (" + ex.FileName + ")", Color.Red);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void velgLagerViewpointDato()
        {
            try
            {
                if (EmptyStoreDatabase())
                    return;

                var datovelger = new VelgDato(appConfig);
                if (datovelger.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SaveSettings();
                    DateTime dato = obsolete.FindNearestDate(appConfig.dbStoreViewpoint);
                    if (dato.Date == appConfig.dbStoreViewpoint.Date)
                        Logg.Log("Lager-utvikling følges fra dato " + dato.ToShortDateString(), Color.Green);
                    else
                    {
                        appConfig.dbStoreViewpoint = dato;
                        Logg.Log("Lager-utvikling følges fra dato " + dato.ToShortDateString() + " (valgt nærmeste dato)", Color.Green);
                    }
                    ClearHashStore();
                    RunStore("Obsolete");
                }
                else
                    Logg.Log("Ingen dato valgt.");
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void ClearHashStore(string katArg = "", bool bg = false)
        {
            if (!String.IsNullOrEmpty(katArg))
            {
                if (katArg.Equals("Obsolete"))
                    appConfig.strObsolete = "";
                else if (katArg.Equals("ObsoleteList"))
                    appConfig.strObsoleteList = "";
                else if (katArg.Equals("LagerUkeAnnonser"))
                    appConfig.strLagerWeekly = "";
                else if (katArg.Equals("LagerUkeAnnonserOversikt"))
                    appConfig.strLagerWeeklyOverview = "";
                else if (katArg.Equals("LagerPrisguide"))
                    appConfig.strLagerPrisguide = "";
                else if (katArg.Equals("LagerPrisguideOversikt"))
                    appConfig.strLagerPrisguideOverview = "";
            }
            else
            {
                appConfig.strObsolete = "";
                appConfig.strObsoleteList = "";
                appConfig.strObsoleteImports = "";
                appConfig.strLagerWeekly = "";
                appConfig.strLagerWeeklyOverview = "";
                appConfig.strLagerPrisguide = "";
                appConfig.strLagerPrisguideOverview = "";
            }
            SaveSettings();
            if (!bg && !EmptyStoreDatabase())
            {
                buttonOppdaterLager.BackColor = Color.LightGreen;
                buttonOppdaterLager.ForeColor = SystemColors.ControlText;
            }
        }

        public void ReloadStore(bool forced = false)
        {
            try
            {
                if (forced)
                    ClearHashStore();
                if (EmptyStoreDatabase())
                {
                    Logg.Log("Lager databasen er tom. Importer wobsolete fra Elguide!");
                    webLager.Navigate(htmlImportStore);

                    buttonOppdaterLager.BackColor = SystemColors.ControlLight;
                    buttonOppdaterLager.ForeColor = SystemColors.ControlText;

                    labelStoreDato.Text = "(tom)";
                    labelStoreDatoUnder.Text = "";
                    labelStoreDato.ForeColor = SystemColors.ControlText;
                    labelStoreDatoUnder.ForeColor = SystemColors.ControlText;

                    ShowHideGui_EmptyStore(false);
                }
                else
                {
                    ShowHideGui_EmptyStore(true);

                    if (!autoMode)
                        UpdateStore();
                    moveStoreDate(0, true);

                    labelStoreDato.Text = appConfig.dbStoreTo.ToString("dddd", norway);
                    labelStoreDatoUnder.Text = appConfig.dbStoreTo.ToString("d. MMM", norway);

                    if ((DateTime.Now - appConfig.dbStoreTo).Days >= 3)
                    {
                        labelStoreDato.ForeColor = Color.Red;
                        labelStoreDatoUnder.ForeColor = Color.Red;
                    }
                    else
                    {
                        labelStoreDato.ForeColor = SystemColors.ControlText;
                        labelStoreDatoUnder.ForeColor = SystemColors.ControlText;
                    }
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Kritisk feil ved initialisering av lager databasen.\nInstaller programmet på nytt hvis problemet vedvarer.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void moveStoreDate(int m = 0, bool reload = false)
        {
            try
            {
                if (!EmptyStoreDatabase())
                {
                    var d = pickerLagerDato.Value;
                    if (m == 1) // gå tilbake en måned
                    {
                        if (appConfig.dbStoreFrom.Date <= d.AddMonths(-1))
                            pickerLagerDato.Value = d.AddMonths(-1);
                        else
                            pickerLagerDato.Value = appConfig.dbStoreFrom;
                    }
                    if (m == 2) // gå tilbake en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbStoreFrom.Date <= d.AddDays(-1) && d.AddDays(-1).DayOfWeek != DayOfWeek.Sunday)
                                pickerLagerDato.Value = d.AddDays(-1);
                            if (appConfig.dbStoreFrom.Date <= d.AddDays(-2) && d.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
                                pickerLagerDato.Value = d.AddDays(-2);
                        }
                        else
                        {
                            if (appConfig.dbStoreFrom.Date <= d.AddDays(-1))
                                pickerLagerDato.Value = d.AddDays(-1);
                        }
                    }
                    if (m == 3) // gå fram en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbStoreTo.Date >= d.AddDays(1) && d.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                                pickerLagerDato.Value = d.AddDays(1);
                            if (appConfig.dbStoreTo.Date >= d.AddDays(2) && d.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                                pickerLagerDato.Value = d.AddDays(2);
                        }
                        else
                        {
                            if (appConfig.dbStoreTo.Date >= d.AddDays(1))
                                pickerLagerDato.Value = d.AddDays(1);
                        }
                    }
                    if (m == 4) // gå fram en måned
                    {
                        if (appConfig.dbStoreTo.Date >= d.AddMonths(1))
                            pickerLagerDato.Value = d.AddMonths(1);
                        else
                            pickerLagerDato.Value = appConfig.dbStoreTo;
                    }
                    d = pickerLagerDato.Value;
                    if (d.Date >= appConfig.dbStoreTo.Date)
                    {
                        buttonLagerF.Enabled = false; // fremover knapp
                        buttonLagerFF.Enabled = false; // fremover knapp
                    }
                    else
                    {
                        buttonLagerF.Enabled = true; // fremover knapp
                        buttonLagerFF.Enabled = true; // fremover knapp
                    }
                    if (d.Date <= appConfig.dbStoreFrom.Date)
                    {
                        buttonLagerBF.Enabled = false; // bakover knapp
                        buttonLagerB.Enabled = false; // bakover knapp
                    }
                    else
                    {
                        buttonLagerBF.Enabled = true; // bakover knapp
                        buttonLagerB.Enabled = true; // bakover knapp
                    }

                    if (Loaded && reload)
                    {
                        if (!IsBusy())
                            UpdateStore();
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }
        public void UpdateStore(string katArg = "")
        {
            string page = currentPage();
            if (String.IsNullOrEmpty(katArg) && !String.IsNullOrEmpty(page))
                ClearHashStore(page);

            if (!String.IsNullOrEmpty(katArg))
                RunStore(katArg);
            else
            {
                if (!String.IsNullOrEmpty(page))
                    RunStore(page);
                else if (!String.IsNullOrEmpty(savedPage))
                    RunStore(savedStorePage);
                else if (!String.IsNullOrEmpty(appConfig.savedPage))
                    RunStore(appConfig.savedStorePage);
                else
                    RunStore("Obsolete");
            }
        }

        private void RunStore(string katArg)
        {
            if (chkStorePicker != pickerLagerDato.Value && !bwStore.IsBusy && !String.IsNullOrEmpty(katArg))
            {
                HighlightStoreButton(katArg);
                if (!EmptyStoreDatabase())
                {
                    groupBoxLagerKat.Enabled = false;
                    bwStore.RunWorkerAsync(katArg);
                }
                else
                    webLager.Navigate(htmlImportStore);
            }
            chkStorePicker = rangeMin;
        }

        private void HighlightStoreButton(string katArg = "")
        {
            try
            {
                buttonLagerUkuListe.BackColor = SystemColors.ControlLight;
                buttonLagerStatus.BackColor = SystemColors.ControlLight;
                buttonLagerWeekly.BackColor = SystemColors.ControlLight;
                buttonLagerWeeklyOverview.BackColor = SystemColors.ControlLight;
                buttonLagerPrisguide.BackColor = SystemColors.ControlLight;
                buttonLagerPrisguideOverview.BackColor = SystemColors.ControlLight;
                buttonLagerImports.BackColor = SystemColors.ControlLight;

                if (katArg == "Obsolete")
                    buttonLagerStatus.BackColor = Color.LightSkyBlue;
                else if (katArg == "ObsoleteList")
                    buttonLagerUkuListe.BackColor = Color.LightSkyBlue;
                else if (katArg.Equals("LagerUkeAnnonser"))
                    buttonLagerWeekly.BackColor = Color.LightSkyBlue;
                else if (katArg.Equals("LagerUkeAnnonserOversikt"))
                    buttonLagerWeeklyOverview.BackColor = Color.LightSkyBlue;
                else if (katArg.Equals("LagerPrisguide"))
                    buttonLagerPrisguide.BackColor = Color.LightSkyBlue;
                else if (katArg.Equals("LagerPrisguideOversikt"))
                    buttonLagerPrisguideOverview.BackColor = Color.LightSkyBlue;
                else if (katArg == "ObsoleteImports")
                    buttonLagerImports.BackColor = Color.LightSkyBlue;

                this.Update();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwStore_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string katArg = (string)e.Argument;
            string newHash = appConfig.Avdeling + pickerLagerDato.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
            if (katArg == "Obsolete")
            {
                BuildStoreStatus();
                appConfig.strObsolete = newHash;
            }
            else if (katArg == "ObsoleteList")
            {
                BuildStoreObsoleteList();
                appConfig.strObsoleteList = newHash;
            }
            else if (katArg == "ObsoleteImports")
            {
                BuildStoreObsoleteImports();
                appConfig.strObsoleteImports = newHash;
            }
            else if (katArg.Equals("LagerUkeAnnonser"))
            {
                BuildStoreWeekly();
                appConfig.strLagerWeekly = newHash;
            }
            else if (katArg.Equals("LagerUkeAnnonserOversikt"))
            {
                BuildStoreWeeklyOverview();
                appConfig.strLagerWeeklyOverview = newHash;
            }
            else if (katArg.Equals("LagerPrisguide"))
            {
                BuildStorePrisguide();
                appConfig.strLagerPrisguide = newHash;
            }
            else if (katArg.Equals("LagerPrisguideOversikt"))
            {
                BuildStorePrisguideOverview();
                appConfig.strLagerPrisguideOverview = newHash;
            }
            else
                Logg.Log("Ingen kategori valgt.");
        }

        private void bwStore_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Logg.Status("Klar.");
            groupBoxLagerKat.Enabled = true;
        }

        private void NavigateStoreWeb(WebBrowserNavigatingEventArgs e)
        {
            try
            {
                var url = e.Url.OriginalString;
                string page = currentPage();
                if (url.Contains("#ukenytt=") && (page.Equals("LagerUkeAnnonser") || page.Equals("LagerUkeAnnonserOversikt")))
                {
                    DateTime date = database.tableWeekly.GetLatestDate();
                    int index = url.IndexOf("#") + 9;
                    string str = url.Substring(index, url.Length - index);

                    if (str.Equals("list"))
                    {
                        ClearHashStore("LagerUkeAnnonserOversikt");
                        RunStore("LagerUkeAnnonserOversikt");
                    }
                    else
                    {
                        DateTime.TryParseExact(Convert.ToString(str), "dd.MM.yyyy", norway, DateTimeStyles.None, out date);

                        ChangeStoreDateTimePicker(date, rangeMin, rangeMax);

                        this.Update();

                        ClearHashStore("LagerUkeAnnonser");
                        RunStore("LagerUkeAnnonser");
                    }
                }
                else if (url.Contains("#link") && !url.Contains("#linkx"))
                {
                    int index = url.IndexOf("#");
                    bool month = true;
                    if (url.Substring(index + 5, 1) == "d")
                        month = false;
                    var type = url.Substring(index + 6, 1);
                    var data = url.Substring(index + 7, url.Length - index - 7);
                    processing.SetVisible = true;
                    processing.SetText = "Søker..";
                    InitDB();
                    tabControlMain.SelectedTab = tabPageTrans;
                    this.Update();
                    SearchDB(page, month, type, data);
                    processing.SetVisible = false;
                    Logg.Status("Ferdig.");
                }
                else if (url.Contains("#prisguide=") && (page.Equals("LagerPrisguide") || page.Equals("LagerPrisguideOversikt")))
                {
                    DateTime date = database.tablePrisguide.GetLatestDate();
                    int index = url.IndexOf("#") + 11;
                    string str = url.Substring(index, url.Length - index);

                    if (str.Equals("list"))
                    {
                        ClearHashStore("LagerPrisguideOversikt");
                        RunStore("LagerPrisguideOversikt");
                    }
                    else
                    {
                        DateTime.TryParseExact(Convert.ToString(str), "dd.MM.yyyy", norway, DateTimeStyles.None, out date);

                        ChangeStoreDateTimePicker(date, rangeMin, rangeMax);

                        this.Update();

                        ClearHashStore("LagerPrisguide");
                        RunStore("LagerPrisguide");
                    }
                }
                else if (url.Contains("#external=") && (page.Equals("LagerPrisguide") || page.Equals("LagerPrisguideOversikt")))
                {
                    int index = url.IndexOf("#") + 10;
                    if (index > 0)
                    {
                        string externalUrl = url.Substring(index);
                        if (externalUrl.StartsWith("http:") || externalUrl.StartsWith("https:"))
                        {
                            System.Diagnostics.Process.Start(url.Substring(index));
                            e.Cancel = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.ErrorDialog(ex, ex.Message, "Navigerings feil");
            }
        }

        private void BuildStoreStatus(bool bg = false)
        {
            string katArg = "Obsolete";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strObsolete);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;

                    dtPick = obsolete.FindNearestDate(dtPick);
                    dtTil = obsolete.FindNearestDate(dtTil);
                    dtFra = obsolete.FindNearestDate(dtFra);

                    if (!appConfig.storeCompareMtd)
                        dtFra = appConfig.dbStoreViewpoint;

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Lager status (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    doc.Add("<h3>Oversikt utgåtte varer hovedlager (" + dtTil.ToString("dddd d. MMMM yyyy", norway) + ")</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webLager.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml(false));

                    if (appConfig.storeShowStoreTwo)
                    {
                        doc.Add("<h3>Oversikt utgåtte varer B/V lager (" + dtTil.ToString("dddd d. MMMM yyyy", norway) + ")</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webLager.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml(true));
                    }

                    doc.Add("<h3>Utvikling hovedlager (Sammenlignet mot " + dtFra.ToString("dddd d. MMMM yyyy", norway) + ")</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webLager.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtmlUtvikling(false));

                    if (appConfig.storeShowStoreTwo)
                    {
                        doc.Add("<h3>Utvikling B/V lager (Sammenlignet mot " + dtFra.ToString("dddd d. MMMM yyyy", norway) + ")</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webLager.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtmlUtvikling(true));
                    }

                    //obsolete.RetreieveHistoryTable();
                    //doc.Add("<img src='" + obsolete.SaveChartImage(appConfig.graphResX, appConfig.graphResY, "MDA", appConfig.Avdeling) + "' style='width:900px;height:auto;'><br>");

                    //doc.Add("<img src='" + obsolete.SaveChartImage(appConfig.graphResX, appConfig.graphResY, "AudioVideo", appConfig.Avdeling) + "' style='width:900px;height:auto;'><br>");

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStoreObsolete, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStoreObsolete);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStoreObsolete);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStoreObsoleteList(bool bg = false)
        {
            string katArg = "ObsoleteList";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strObsoleteList);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Utgåtte varer (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");
                    doc.Add("<span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                    doc.AddRange(ranking.GetTableHtmlUkurantGrupper());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStoreObsoleteList, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStoreObsoleteList);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStoreObsoleteList);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStoreWeekly(bool bg = false)
        {
            string katArg = "LagerUkeAnnonser";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strLagerWeekly);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<span class='Title'>Ukesnytt lagerstatus (" + avdeling.Get(appConfig.Avdeling) + ")</span><span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                    doc.AddRange(ranking.GetTableHtmlWeekly());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStoreWeekly, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStoreWeekly);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStoreWeekly);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStoreWeeklyOverview(bool bg = false)
        {
            string katArg = "LagerUkeAnnonserOversikt";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strLagerWeeklyOverview);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Ukesnytt lagerstatus (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");
                    doc.Add("<span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                    doc.AddRange(ranking.GetTableHtmlWeeklyList());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStoreWeeklyOverview, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStoreWeeklyOverview);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStoreWeeklyOverview);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStorePrisguide(bool bg = false)
        {
            string katArg = "LagerPrisguide";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strLagerPrisguide);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Prisguide.no (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");
                    doc.Add("<span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                    doc.AddRange(ranking.GetTableHtmPrisguide());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStorePrisguide, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStorePrisguide);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStorePrisguide);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStorePrisguideOverview(bool bg = false)
        {
            string katArg = "LagerPrisguideOversikt";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strLagerPrisguideOverview);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<h1>Prisguide.no (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");
                    doc.Add("<span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                    doc.AddRange(ranking.GetTableHtmPrisguideList());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStorePrisguideOverview, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStorePrisguideOverview);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStorePrisguideOverview);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildStoreObsoleteImports(bool bg = false)
        {
            string katArg = "ObsoleteImports";
            bool abort = HarSisteVersjonStore(katArg, appConfig.strObsoleteImports);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedStorePage = katArg;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webLager.Navigate(htmlLoading);
                    var doc = new List<string>();

                    DateTime dtPick = pickerLagerDato.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingStore(this, dtFra, dtTil, dtPick, obsolete);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true);

                    doc.Add("<span class='Title'>Lager status (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                    doc.Add("<span class='xTitle'>Viser alle importeringer</span>");

                    doc.AddRange(ranking.GetTableHtmlReport());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        webLager.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webLager.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlStoreObsoleteImports, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webLager.Navigate(htmlStoreObsoleteImports);
                            if (!bg) Logg.Log("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webLager.Navigate(htmlStoreObsoleteImports);
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved generering av ranking for [" + katArg + "] Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webLager.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        public bool HarSisteVersjonStore(string katArg, string oldHash)
        {
            try
            {
                if (datoPeriodeVelger)
                    return false;

                string html = "";
                if (katArg == "Obsolete")
                    html = htmlStoreObsolete;
                else if (katArg == "ObsoleteList")
                    html = htmlStoreObsoleteList;
                else if (katArg == "ObsoleteImports")
                    html = htmlStoreObsoleteImports;
                else if (katArg.Equals("LagerUkeAnnonser"))
                    html = htmlStoreWeekly;
                else if (katArg.Equals("LagerUkeAnnonserOversikt"))
                    html = htmlStoreWeeklyOverview;
                else if (katArg.Equals("LagerPrisguide"))
                    html = htmlStorePrisguide;
                else
                    return false;

                if (File.Exists(html))
                {
                    string newHash = appConfig.Avdeling + pickerLagerDato.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                    if (newHash == oldHash)
                    {
                        buttonOppdaterLager.BackColor = SystemColors.ControlLight;
                        buttonOppdaterLager.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: bruker mellomlagret kopi!");
                        return true;
                    }
                    else
                    {
                        buttonOppdaterLager.BackColor = SystemColors.ControlLight;
                        buttonOppdaterLager.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: Genererer ny side..");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
            return false;
        }
    }
}
