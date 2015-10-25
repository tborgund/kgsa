using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    
    partial class FormMain
    {
        delegate void SetCancelButtonCallback(bool value);
        private void SetGroupBoxDatoCallback(bool value)
        {
            if (groupRankingChoices.InvokeRequired)
            {
                SetCancelButtonCallback d = new SetCancelButtonCallback(SetGroupBoxDatoCallback);
                this.Invoke(d, new object[] { value });
            }
            else
                groupRankingChoices.Enabled = value;
        }

        private void RunRanking(string katArg)
        {
            if (chkPicker != pickerRankingDate.Value && !bwRanking.IsBusy && !String.IsNullOrEmpty(katArg))
            {
                HighlightButton(katArg);
                if (!EmptyDatabase())
                    bwRanking.RunWorkerAsync(katArg);
                else
                    webRanking.Navigate(htmlImport);
            }
            chkPicker = rangeMin;
        }

        private void RunVinnSelger(string selgerArg)
        {
            if (!EmptyDatabase())
            {
                groupRankingChoices.Enabled = false;
                Log.n("Oppdaterer [" + selgerArg + "] ..", Color.Black, false, true);
                bwVinnSelger.RunWorkerAsync(selgerArg);
            }
            else
                webRanking.Navigate(htmlImport);

        }

        public void RunBudget(BudgetCategory cat)
        {
            if (chkBudgetPicker != pickerBudget.Value && !bwBudget.IsBusy)
            {
                HighlightBudgetButton(cat);
                if (!EmptyDatabase())
                {
                    groupBudgetChoices.Enabled = false;
                    bwBudget.RunWorkerAsync(cat);
                }
                else
                    webBudget.Navigate(htmlImport);
            }
            chkBudgetPicker = rangeMin;
        }

        private void bwVinnSelger_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string selgerArg = (string)e.Argument;

            BuildVinnRankingSelger(selgerArg);
        }

        private void bwRanking_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string katArg = (string)e.Argument;
            if (katArg == null)
                katArg = "";

            string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
            if (katArg == "Oversikt")
            {
                BuildOversiktRanking();
                appConfig.strOversikt = newHash;
            }
            else if (katArg == "Butikk")
            {
                BuildButikkRanking();
                appConfig.strButikk = newHash;
            }
            else if (katArg == "KnowHow")
            {
                BuildKnowHowRanking();
                appConfig.strKnowHow = newHash;
            }
            else if (katArg == "Data")
            {
                BuildDataRanking();
                appConfig.strData = newHash;
            }
            else if (katArg == "AudioVideo")
            {
                BuildAudioVideoRanking();
                appConfig.strAudioVideo = newHash;
            }
            else if (katArg == "Tele")
            {
                BuildTeleRanking();
                appConfig.strTele = newHash;
            }
            else if (katArg == "Toppselgere")
            {
                BuildToppselgereRanking();
                appConfig.strToppselgere = newHash;
            }
            else if (katArg == "Lister")
            {
                BuildListerRanking();
                appConfig.strLister = newHash;
            }
            else if (katArg == "Vinnprodukter")
            {
                BuildVinnRanking();
                appConfig.strVinnprodukter = newHash;
            }
            else if (katArg == "Tjenester")
            {
                BuildAvdTjenester();
                appConfig.strTjenester = newHash;
            }
            else if (katArg == "Snittpriser")
            {
                BuildAvdSnittpriser();
                appConfig.strSnittpriser = newHash;
            }
            else
                Log.n("Ingen kategori valgt for beregning av ranking.");
        }

        private void bwRanking_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Log.Status("Klar.");
            groupRankingChoices.Enabled = true;
        }

        private void bwBudget_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            BudgetCategory cat = (BudgetCategory)e.Argument;
            MakeBudgetPage(cat, bwBudget);
        }

        public void MakeBudgetPage(BudgetCategory cat, BackgroundWorker bw = null)
        {
            string newHash = appConfig.Avdeling + pickerBudget.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
            if (cat == BudgetCategory.MDA)
            {
                BuildBudget(cat, appConfig.strBudgetMda, htmlBudgetMda);
                appConfig.strBudgetMda = newHash;
            }
            else if (cat == BudgetCategory.AudioVideo)
            {
                BuildBudget(cat, appConfig.strBudgetAv, htmlBudgetAudioVideo);
                appConfig.strBudgetAv = newHash;
            }
            else if (cat == BudgetCategory.SDA)
            {
                BuildBudget(cat, appConfig.strBudgetSda, htmlBudgetSda);
                appConfig.strBudgetSda = newHash;
            }
            else if (cat == BudgetCategory.Tele)
            {
                BuildBudget(cat, appConfig.strBudgetTele, htmlBudgetTele);
                appConfig.strBudgetTele = newHash;
            }
            else if (cat == BudgetCategory.Data)
            {
                BuildBudget(cat, appConfig.strBudgetData, htmlBudgetData);
                appConfig.strBudgetData = newHash;
            }
            else if (cat == BudgetCategory.Cross)
            {
                BuildBudget(cat, appConfig.strBudgetCross, htmlBudgetCross);
                appConfig.strBudgetCross = newHash;
            }
            else if (cat == BudgetCategory.Kasse)
            {
                BuildBudget(cat, appConfig.strBudgetKasse, htmlBudgetKasse);
                appConfig.strBudgetKasse = newHash;
            }
            else if (cat == BudgetCategory.Aftersales)
            {
                BuildBudget(cat, appConfig.strBudgetAftersales, htmlBudgetAftersales);
                appConfig.strBudgetAftersales = newHash;
            }
            else if (cat == BudgetCategory.MDASDA)
            {
                BuildBudget(cat, appConfig.strBudgetMdasda, htmlBudgetMdasda);
                appConfig.strBudgetMdasda = newHash;
            }
            else if (cat == BudgetCategory.Butikk)
            {
                BuildBudget(cat, appConfig.strBudgetButikk, htmlBudgetButikk);
                appConfig.strBudgetButikk = newHash;
            }
            else if (cat == BudgetCategory.Daglig)
            {
                PageBudgetDaily page = new PageBudgetDaily(this, false, bw, webBudget);
                page.BuildPage(cat, appConfig.strBudgetDaily, htmlBudgetDaily, pickerBudget.Value);
                appConfig.strBudgetDaily = newHash;
            }
            else if (cat == BudgetCategory.AlleSelgere)
            {
                PageBudgetAllSales page = new PageBudgetAllSales(this, false, bw, webBudget);
                page.BuildPage(cat, appConfig.strBudgetAllSales, htmlBudgetAllSales, pickerBudget.Value);
                appConfig.strBudgetAllSales = newHash;
            }
            else
                Log.n("Ingen kategori valgt for beregning av budsjett.");
        }

        private void bwBudget_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Log.Status("Klar.");
            groupBudgetChoices.Enabled = true;
        }

        private void ViewGraph(string argKat, bool bg = false, BackgroundWorker bw = null)
        {
            Log.n("Lager [" + argKat + "] grafer..");
            try
            {
                if (!bg)
                {
                    timewatch.Start();
                    webRanking.Navigate(htmlLoading);
                }
                var doc = new List<string>();

                DateTime dtPick = pickerRankingDate.Value;
                DateTime dtFra = dtPick; DateTime dtTil = dtPick;
                if (datoPeriodeVelger && !bg)
                {
                    dtFra = datoPeriodeFra;
                    dtTil = datoPeriodeTil;
                }

                var ytdDate = new DateTime(dtPick.Year, 5, 1);
                if (ytdDate > dtPick)
                    ytdDate = new DateTime(dtPick.AddYears(-1).Year, 5, 1);
                int ytdDays = (dtPick - ytdDate).Days;

                GetHtmlStart(doc, true, "Graf: " + argKat);

                doc.Add("<h1>Historisk " + argKat + " (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                if (!datoPeriodeVelger || bg)
                    doc.Add("<h3>Siste " + appConfig.graphDager + " dager</h3>");
                else
                    doc.Add("<h3>Graf for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                doc.Add("<span class='Loading'>Henter graf..</span>");
                if (!bg)
                    webRanking.DocumentText = string.Join(null, doc.ToArray());
                doc.RemoveAt(doc.Count - 1);

                if (datoPeriodeVelger && !bg)
                    gc.SaveImageChunk(argKat, settingsPath + @"\graphGraf1" + argKat + ".png", appConfig.graphResX, (int)(appConfig.graphResY * 1.5D), "", dtFra, dtTil, false, bw);
                else
                    gc.SaveImageChunk(argKat, settingsPath + @"\graphGraf1" + argKat + ".png", appConfig.graphResX, (int)(appConfig.graphResY * 1.5D),"", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);
                doc.Add("<img src='graphGraf1" + argKat + ".png' class='image' style='width:" + appConfig.graphWidth + "px;'>");

                if (appConfig.graphExtra && !datoPeriodeVelger)
                {
                    doc.Add("<h3>Siste " + (appConfig.graphDager * 3) + " dager</h3>");
                    doc.Add("<span class='Loading'>Henter graf..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                    {
                        File.WriteAllLines(htmlGraf, doc.ToArray(), Encoding.Unicode);
                        webRanking.Navigate(htmlGraf);
                    }
                    doc.RemoveAt(doc.Count - 1);

                    gc.SaveImageChunk(argKat, settingsPath + @"\graphGraf2" + argKat + ".png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-(appConfig.graphDager * 3)), dtTil, false, bw);
                    doc.Add("<img src='graphGraf2" + argKat + ".png' class='image' style='width:" + appConfig.graphWidth + "px;'>");
                }

                if (!datoPeriodeVelger)
                {
                    doc.Add("<h3>Year To Date</h3>");
                    doc.Add("<span class='Loading'>Henter graf..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                    {
                        File.WriteAllLines(htmlGraf, doc.ToArray(), Encoding.Unicode);
                        webRanking.Navigate(htmlGraf);
                    }
                    doc.RemoveAt(doc.Count - 1);

                    gc.SaveImageChunk(argKat, settingsPath + @"\graphGrafYTD" + argKat + ".png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-ytdDays), dtTil, false, bw);
                    doc.Add("<img src='graphGrafYTD" + argKat + ".png' class='image' style='width:" + appConfig.graphWidth + "px;'>");
                }

                doc.Add(Resources.htmlEnd);
                if (!bg)
                    File.WriteAllLines(htmlGraf, doc.ToArray(), Encoding.Unicode);
                else
                    File.WriteAllLines(settingsPath + @"\rankingGraf" + argKat + ".html", doc.ToArray(), Encoding.Unicode);

                if (!bg)
                {
                    webRanking.Navigate(htmlGraf);
                    Log.n("Generering av graf for " + argKat + " tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch(Exception ex)
            {
                Log.n("Feil ved generering av grafikk for " + argKat + ". Exception: " + ex.ToString(), Color.Red);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av graf for " + argKat, ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        public static bool stopRanking = false;
        private void BuildButikkRanking(bool bg = false, BackgroundWorker bw = null)
        {
            Log.n("Lager [Butikk] ranking..");
            bool abort = HarSisteVersjon("Butikk", appConfig.strButikk);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = "Butikk";
                if (!abort)
                {
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingButikk(this, dtFra, dtTil, dtPick);

                    GetHtmlStart(doc, true, "Ranking: Butikk");

                    doc.Add("<h1>Oversikt butikk (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    openXml.DeleteDocument("Butikk", dtPick);

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Dagens tall for " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Tall for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml("dag"));

                    if (!datoPeriodeVelger || bg)
                    {
                        doc.Add("<h3>" + StringRankingDato(dtPick) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml("måned"));

                        if (appConfig.rankingCompareLastyear > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddYears(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareHtml());
                        }
                        if (appConfig.rankingCompareLastmonth > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddMonths(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                        }
                        if (appConfig.rankingShowLastWeek)
                        {
                            doc.Add("<h3>Sist uke - Uke " + database.GetIso8601WeekOfYear(dtPick) + " (Fra " + database.GetStartOfLastWholeWeek(dtPick).ToString("d. MMM") + " til " + database.GetStartOfLastWholeWeek(dtPick).AddDays(6).ToString("d. MMM") + ")</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableLastWholeWeek());
                        }
                        if (appConfig.favVis && Favoritter.Count > 0)
                        {
                            doc.Add("<h3>Favoritt avdelinger " + StringRankingDato(dtPick) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetAvdHtml());
                        }
                    }
                    if (appConfig.graphVis && appConfig.graphButikk && appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("Butikk", settingsPath + @"\graphButikk.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil, false, bw);
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager</h3>"); 
                        }
                        else
                        {
                            gc.SaveImageChunk("Butikk", settingsPath + @"\graphButikk.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager</h3>");
                        }
                        doc.Add("<a href='#grafButikk'><img src='graphButikk.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");

                        //doc.AddRange(gc.SaveGraphData());
                    }

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash("Butikk");
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingButikk, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingButikk);
                            if (!bg) Log.n("Ranking [Butikk] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingButikk);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [Butikk]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        private void BuildBudget(BudgetCategory cat, string currentHash, string htmlPage, bool bg = false)
        {
            string katArg = BudgetCategoryClass.TypeToName(cat);
            bool abort = HarSisteVersjonBudget(cat, currentHash);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedBudgetPage = cat;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + " Budget]..");
                    if (!bg)
                        webBudget.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerBudget.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingBudget(this, dtFra, dtTil, cat);

                    if (ranking.budgetInfo == null)
                    {
                        ClearBudgetHash(cat);
                        Log.n("Mangler budsjett for " + katArg + ".");
                        doc = new List<string>();
                        doc.Add(File.ReadAllText(htmlSetupBudget));
                        File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                        if (!bg)
                            webBudget.Navigate(htmlPage);
                        return;
                    }

                    GetHtmlStart(doc, true, "Ranking: Budsjett");

                    doc.Add("<h1>Budsjett avdeling: " + katArg + " (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    doc.Add("<table style='width:100%;display:none;' class='hidePdf'><tr><td>Hopp til: ");
                    if (ranking.budgetInfo.acc > 0)
                        doc.Add("<span class=\"link_button\"><a href='#Tilbehør'>Tilbehør</a></span>&nbsp;");
                    if (ranking.budgetInfo.finans > 0)
                        doc.Add("<span class=\"link_button\"><a href='#Finansiering'>Finansiering</a></span>&nbsp;");
                    if (ranking.budgetInfo.rtgsa > 0)
                        if (cat != BudgetCategory.MDA && cat != BudgetCategory.SDA)
                            doc.Add("<span class=\"link_button\"><a href='#RTG/SA'>RTG/SA</a></span>&nbsp;");
                    if (ranking.budgetInfo.strom > 0)
                        doc.Add("<span class=\"link_button\"><a href='#Norges Energi'>Strøm</a></span>&nbsp;");
                    if (ranking.budgetInfo.ta > 0)
                        doc.Add("<span class=\"link_button\"><a href='#Trygghetsavtale'>Trygghetsavtale</a></span>&nbsp;");
                    if (ranking.budgetInfo.vinn > 0)
                        doc.Add("<span class=\"link_button\"><a href='#Vinnprodukt'>Vinnprodukter</a></span>&nbsp;");
                    doc.Add("</td></tr></table>");

                    if (ranking.budgetInfo.inntjening > 0)
                        doc.AddRange(ranking.GetTableHtml(cat));

                    if (ranking.budgetInfo.acc > 0)
                        doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.Acc));

                    if (ranking.budgetInfo.finans > 0)
                        doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.Finans));

                    if (ranking.budgetInfo.rtgsa > 0)
                        if (cat != BudgetCategory.MDA && cat != BudgetCategory.SDA)
                            doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.Rtgsa));

                    if (ranking.budgetInfo.strom > 0)
                        doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.Strom));

                    if (ranking.budgetInfo.ta > 0)
                        doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.TA));

                    if (ranking.budgetInfo.vinn > 0)
                        doc.AddRange(ranking.GetTableHtmlProduct(cat, BudgetType.Vinnprodukt));

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearBudgetHash(cat);
                        Log.n("Ranking stoppet.", Color.Red);
                        webBudget.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webBudget.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webBudget.Navigate(htmlPage);
                            if (!bg) Log.n("Budsjett [" + katArg + " Budget] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webBudget.Navigate(htmlPage);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webBudget.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av budsjett for [" + katArg + " Budget]", ex);
                    errorMsg.ShowDialog(this);
                }
            }
        }

        public void BuildWebRankingData(List<string> doc, DateTime date)
        {
            Log.n("Web: Lager [Data] ranking..", Color.DarkGoldenrod);
            try
            {
                DateTime dtFra = GetFirstDayOfMonth(date);
                DateTime dtTil = date;
                timewatch.Start();

                var ranking = new RankingData(this, dtFra, dtTil, dtTil);

                GetHtmlStart(doc, true, "Ranking: Data");

                doc.Add("<h1>Avdeling: Data (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                doc.Add("<h2>Dagens ranking for " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("dag"));

                doc.Add("<h2>" + StringRankingDato(dtTil) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("måned"));

                if (appConfig.rankingCompareLastyear > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddYears(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareHtml());
                }
                if (appConfig.rankingCompareLastmonth > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddMonths(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                }
                if (appConfig.favVis && Favoritter.Count > 0)
                {
                    doc.Add("<h2>Favoritt avdelinger " + StringRankingDato(dtTil) + "</h2>");
                    doc.AddRange(ranking.GetAvdHtml());
                }

                if (appConfig.graphVis && appConfig.graphData)
                {
                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("<div class='Hidegraph'>");

                    doc.Add("<h2>Siste " + appConfig.graphDager + " dager Datamaskiner</h2>");
                    gc.SaveImageChunk("Data", settingsPath + @"\graphDataWeb.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil);
                    doc.Add("<a href='#grafData'><img src='graphDataWeb.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                    doc.Add("<h2>Siste " + appConfig.graphDager + " dager Nettbrett:</h2>");
                    gc.SaveImageChunk("Nettbrett", settingsPath + @"\graphNettbrettWeb.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, null);
                    doc.Add("<a href='#grafNettbrett'><img src='graphNettbrettWeb.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("</div>");
                }

                doc.Add(Resources.htmlEnd);
                Log.n("Web: Ranking [Data] tok " + timewatch.Stop() + " sekunder.", Color.DarkGoldenrod, true);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void BuildWebRankingAudioVideo(List<string> doc, DateTime date)
        {
            Log.n("Web: Lager [AudioVideo] ranking..", Color.DarkGoldenrod);
            try
            {
                DateTime dtFra = GetFirstDayOfMonth(date);
                DateTime dtTil = date;
                timewatch.Start();

                var ranking = new RankingAudioVideo(this, dtFra, dtTil, dtTil);

                GetHtmlStart(doc, true, "Ranking: Lyd & Bilde");

                doc.Add("<h1>Avdeling: Lyd & Bilde (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                doc.Add("<h2>Dagens ranking for " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("dag"));

                doc.Add("<h2>" + StringRankingDato(dtTil) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("måned"));

                if (appConfig.rankingCompareLastyear > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddYears(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareHtml());
                }
                if (appConfig.rankingCompareLastmonth > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddMonths(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                }
                if (appConfig.favVis && Favoritter.Count > 0)
                {
                    doc.Add("<h2>Favoritt avdelinger " + StringRankingDato(dtTil) + "</h2>");
                    doc.AddRange(ranking.GetAvdHtml());
                }

                if (appConfig.graphVis && appConfig.graphAudioVideo)
                {
                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("<div class='Hidegraph'>");

                    doc.Add("<h2>Siste " + appConfig.graphDager + " dager Datamaskiner</h2>");
                    gc.SaveImageChunk("AudioVideo", settingsPath + @"\graphAudioVideoWeb.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, null);
                    doc.Add("<a href='#grafAudioVideo'><img src='graphAudioVideoWeb.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("</div>");
                }

                doc.Add(Resources.htmlEnd);
                Log.n("Web: Ranking [AudioVideo] tok " + timewatch.Stop() + " sekunder.", Color.DarkGoldenrod, true);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        public void BuildWebRankingTele(List<string> doc, DateTime date)
        {
            Log.n("Web: Lager [Tele] ranking..", Color.DarkGoldenrod);
            try
            {
                DateTime dtFra = GetFirstDayOfMonth(date);
                DateTime dtTil = date;
                timewatch.Start();

                var ranking = new RankingTele(this, dtFra, dtTil, dtTil);

                GetHtmlStart(doc, true, "Ranking: Tele");

                doc.Add("<h1>Avdeling: Tele (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                doc.Add("<h2>Dagens ranking for " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("dag"));

                doc.Add("<h2>" + StringRankingDato(dtTil) + "</h2>");
                doc.AddRange(ranking.GetTableHtml("måned"));

                if (appConfig.rankingCompareLastyear > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddYears(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareHtml());
                }
                if (appConfig.rankingCompareLastmonth > 0)
                {
                    doc.Add("<h2>Sammenlignet " + StringRankingDato(dtTil.AddMonths(-1)) + "</h2>");
                    doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                }
                if (appConfig.favVis && Favoritter.Count > 0)
                {
                    doc.Add("<h2>Favoritt avdelinger " + StringRankingDato(dtTil) + "</h2>");
                    doc.AddRange(ranking.GetAvdHtml());
                }

                if (appConfig.graphVis && appConfig.graphTele)
                {
                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("<div class='Hidegraph'>");

                    doc.Add("<h2>Siste " + appConfig.graphDager + " dager Datamaskiner:</h2>");
                    gc.SaveImageChunk("Tele", settingsPath + @"\graphTeleWeb.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, null);
                    doc.Add("<a href='#grafTele'><img src='graphTeleWeb.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                    if (appConfig.pdfExpandedGraphs)
                        doc.Add("</div>");
                }

                doc.Add(Resources.htmlEnd);
                Log.n("Web: Ranking [Tele] tok " + timewatch.Stop() + " sekunder.", Color.DarkGoldenrod, true);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void BuildDataRanking(bool bg = false, BackgroundWorker bw = null)
        {
            string katArg = "Data";
            bool abort = HarSisteVersjon(katArg, appConfig.strData);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingData(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>Avdeling: Data (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Dagens ranking for " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Ranking for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml("dag"));

                    if (!datoPeriodeVelger || bg)
                    {
                        doc.Add("<h3>" + StringRankingDato(dtPick) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml("måned"));

                        if (appConfig.rankingCompareLastyear > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddYears(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareHtml());
                        }
                        if (appConfig.rankingCompareLastmonth > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddMonths(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                        }
                        if (appConfig.favVis && Favoritter.Count > 0)
                        {
                            doc.Add("<h3>Favoritt avdelinger " + StringRankingDato(dtPick) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetAvdHtml());
                        }
                    }
                    if (appConfig.graphVis && appConfig.graphData)
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager Datamaskiner</h3>");
                        else
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager Datamaskiner</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("Data", settingsPath + @"\graphDataPeriode.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil);
                            doc.Add("<a href='#grafData'><img src='graphDataPeriode.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }
                        else
                        {
                            gc.SaveImageChunk("Data", settingsPath + @"\graphData.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil);
                            doc.Add("<a href='#grafData'><img src='graphData.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h2>Siste " + appConfig.graphDager + " dager Nettbrett</h3>");
                        else
                            doc.Add("<h2>Viser " + (dtTil - dtFra).Days + " dager Nettbrett</h3>");

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("Nettbrett", settingsPath + @"\graphNettbrettPeriode.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil, false, bw);
                            doc.Add("<a href='#grafNettbrett'><img src='graphNettbrettPeriode.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }
                        else
                        {
                            gc.SaveImageChunk("Nettbrett", settingsPath + @"\graphNettbrett.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);
                            doc.Add("<a href='#grafNettbrett'><img src='graphNettbrett.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");
                    }


                    doc.Add(Resources.htmlEnd);
                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingData, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingData);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingData);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildKnowHowRanking(bool bg = false)
        {
            string katArg = "KnowHow";
            bool abort = HarSisteVersjon("KnowHow", appConfig.strKnowHow);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingKnowHow(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>Avdeling: KnowHow (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Dagens ranking for " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Ranking for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml("dag"));

                    if (!datoPeriodeVelger || bg)
                    {
                        doc.Add("<h3>" + StringRankingDato(dtPick) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml("måned"));

                        if (appConfig.rankingCompareLastyear > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddYears(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareHtml());
                        }
                        if (appConfig.rankingCompareLastmonth > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddMonths(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                        }
                        if (appConfig.favVis && Favoritter.Count > 0)
                        {
                            doc.Add("<h3>Favoritt avdelinger " + StringRankingDato(dtPick) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetAvdHtml());
                        }
                    }
                    if (appConfig.graphVis && appConfig.graphKnowHow)
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager</span>");
                        else
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager</span>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("KnowHow", settingsPath + @"\graphKnowHowPeriode.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil);
                            doc.Add("<a href='#grafKnowHow'><img src='graphKnowHowPeriode.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }
                        else
                        {
                            gc.SaveImageChunk("KnowHow", settingsPath + @"\graphKnowHow.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil);
                            doc.Add("<a href='#grafKnowHow'><img src='graphKnowHow.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");
                    }

                    doc.Add(Resources.htmlEnd);
                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingKnowHow, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingKnowHow);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingKnowHow);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildAudioVideoRanking(bool bg = false, BackgroundWorker bw = null)
        {
            string katArg = "AudioVideo";
            bool abort = HarSisteVersjon(katArg, appConfig.strAudioVideo);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingAudioVideo(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: Lyd & Bilde");

                    doc.Add("<h1>Avdeling: Lyd & Bilde (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Dagens ranking for " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Ranking for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml("dag"));

                    if (!datoPeriodeVelger || bg)
                    {
                        doc.Add("<h3>" + StringRankingDato(dtPick) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml("måned"));

                        if (appConfig.rankingCompareLastyear > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddYears(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareHtml());
                        }
                        if (appConfig.rankingCompareLastmonth > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddMonths(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                        }
                        if (appConfig.favVis && Favoritter.Count > 0)
                        {
                            doc.Add("<h3>Favoritt avdelinger " + StringRankingDato(dtPick) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetAvdHtml());
                        }
                    }
                    if (appConfig.graphVis && appConfig.graphAudioVideo)
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager</h3>");
                        else
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager</h3>");

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("AudioVideo", settingsPath + @"\graphAudioVideoPeriode.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil, false, bw);
                            doc.Add("<a href='#grafAudioVideo'><img src='graphAudioVideoPeriode.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }
                        else
                        {
                            gc.SaveImageChunk("AudioVideo", settingsPath + @"\graphAudioVideo.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);
                            doc.Add("<a href='#grafAudioVideo'><img src='graphAudioVideo.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");
                    }

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingAudioVideo, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingAudioVideo);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingAudioVideo);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildTeleRanking(bool bg = false, BackgroundWorker bw = null)
        {
            string katArg = "Tele";
            bool abort = HarSisteVersjon(katArg, appConfig.strTele);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingTele(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>Avdeling: Tele (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Dagens ranking for " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Ranking for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtml("dag"));

                    if (!datoPeriodeVelger || bg)
                    {
                        doc.Add("<h3>" + StringRankingDato(dtPick) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtml("måned"));

                        if (appConfig.rankingCompareLastyear > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddYears(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareHtml());
                        }
                        if (appConfig.rankingCompareLastmonth > 0)
                        {
                            doc.Add("<h3>" + StringRankingDato(dtPick.AddMonths(-1)) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetTableCompareLastMonthHtml());
                        }
                        if (appConfig.favVis && Favoritter.Count > 0)
                        {
                            doc.Add("<h3>Favoritt avdelinger " + StringRankingDato(dtPick) + "</h3>");
                            doc.Add("<span class='Loading'>Beregner..</span>");
                            if (!bg && timewatch.ReadyForRefresh())
                                webRanking.DocumentText = string.Join(null, doc.ToArray());
                            doc.RemoveAt(doc.Count - 1);
                            doc.AddRange(ranking.GetAvdHtml());
                        }
                    }
                    if (appConfig.graphVis && appConfig.graphTele)
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager</h3>");
                        else
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager</h3>");

                        if (datoPeriodeVelger && !bg)
                        {
                            gc.SaveImageChunk("Tele", settingsPath + @"\graphTelePeriode.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil, false, bw);
                            doc.Add("<a href='#grafTele'><img src='graphTelePeriode.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }
                        else
                        {
                            gc.SaveImageChunk("Tele", settingsPath + @"\graphTele.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);
                            doc.Add("<a href='#grafTele'><img src='graphTele.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");
                        }

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");
                    }

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg) 
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingTele, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingTele);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingTele);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildOversiktRanking(bool bg = false, BackgroundWorker bw = null)
        {
            string katArg = "Oversikt";
            bool abort = HarSisteVersjon(katArg, appConfig.strOversikt);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (appConfig.dbTo.Month != dtPick.Month && appConfig.dbTo.Year != dtPick.Year) // Hent hele måneden hvis vi IKKE er i siste måneden
                        dtTil = GetLastDayOfMonth(dtPick);
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingOversikt(this, dtFra, dtTil);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>Oversikt selgere (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Detaljert oversikt (Del 1) " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                    else
                        doc.Add("<h3>Detaljert oversikt (Del 1) for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.AddRange(ranking.GetTableHtmlPrimary());

                    if (appConfig.importSetting.StartsWith("Full"))
                    {
                        doc.Add("<div class='no-break'>");
                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Detaljert oversikt (Del 2) " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                        else
                            doc.Add("<h3>Detaljert oversikt (Del 2) for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                        doc.Add("<br><span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtmlSecondary());
                        doc.Add("</div>");
                    }

                    if (appConfig.graphVis && appConfig.graphOversikt)
                    {
                        doc.Add("<span class='Loading'>Henter graf..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("<div class='Hidegraph'>");

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h3>Siste " + appConfig.graphDager + " dager</h3>");
                        else
                            doc.Add("<h3>Viser " + (datoPeriodeTil - datoPeriodeFra).Days + " dager</h3>");

                        if (datoPeriodeVelger && !bg)
                            gc.SaveImageChunk(katArg, settingsPath + @"\graphOversikt.png", appConfig.graphResX, appConfig.graphResY, "", dtFra, dtTil, false, bw);
                        else
                            gc.SaveImageChunk(katArg, settingsPath + @"\graphOversikt.png", appConfig.graphResX, appConfig.graphResY, "", dtTil.AddDays(-appConfig.graphDager), dtTil, false, bw);

                        doc.Add("<a href='#grafOversikt'><img src='graphOversikt.png' style='width:" + appConfig.graphWidth + "px;height:auto;'></a>");

                        if (appConfig.pdfExpandedGraphs)
                            doc.Add("</div>");
                    }

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingOversikt, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingOversikt);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingOversikt);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildToppselgereRanking(bool bg = false)
        {
            string katArg = "Toppselgere";
            bool abort = HarSisteVersjon(katArg, appConfig.strToppselgere);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (appConfig.dbTo.Month != dtPick.Month && appConfig.dbTo.Year != dtPick.Year) // Hent hele måneden hvis vi IKKE er i siste måneden
                        dtTil = GetLastDayOfMonth(dtPick);
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingToppselgere(this, dtFra, dtTil);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>" + katArg + " (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (appConfig.bestofVisBesteLastOpenday && !datoPeriodeVelger)
                    {
                        doc.Add("<div class='no-break'>");
                        doc.Add("<h3>Beste selgere sist åpningsdag " + dtTil.ToString("dddd d. MMMM", norway) + "</h3>");
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg)
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.Add("<table style='width:100%'><tr><td>");
                        doc.AddRange(ranking.GetToppListAllLastOpenDay(dtTil));
                        doc.Add("</td></tr></table>");
                        doc.Add("</div>");
                    }

                    doc.Add("<div class='no-break'>");
                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Beste selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                    else
                        doc.Add("<h3>Beste selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg)
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);
                    doc.Add("<table style='width:100%'><tr><td>");
                    doc.AddRange(ranking.GetToppListAll(dtTil));
                    if (appConfig.bestofHoppoverKasse)
                        doc.Add("<br><span class='information'>* = Viser ikke selgere fra 'Kasse'-avdeling.</span>");
                    doc.Add("</td></tr></table>");
                    doc.Add("</div>");

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingToppselgere, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingToppselgere);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingToppselgere);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildListerRanking(bool bg = false)
        {
            string katArg = "Lister";
            bool abort = HarSisteVersjon(katArg, appConfig.strLister);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (appConfig.dbTo.Month != dtPick.Month && appConfig.dbTo.Year != dtPick.Year) // Hent hele måneden hvis vi IKKE er i siste måneden
                        dtTil = GetLastDayOfMonth(dtPick);
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingToppselgere(this, dtFra, dtTil);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, false, "Ranking: " + katArg); 

                    doc.Add("<h1>Lister (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    ranking.GetToppListLarge(dtTil, null, true); // Generer lister

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h2>Alle salg RTG/SA " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                    else
                        doc.Add("<h2>Alle salg RTG/SA for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                    doc.AddRange(ranking.GetToppListLarge(dtTil, "RTGSA"));

                    if (appConfig.importSetting.StartsWith("Full"))
                    {
                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h2>Alle salg Trygghetsavtale " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                        else
                            doc.Add("<h2>Alle salg Trygghetsavtale for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                        doc.AddRange(ranking.GetToppListLarge(dtTil, "TA"));

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h2>Alle salg Finansiering " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                        else
                            doc.Add("<h2>Alle salg Finansiering for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                        doc.AddRange(ranking.GetToppListLarge(dtTil, "Finans"));

                        if (!datoPeriodeVelger || bg)
                            doc.Add("<h2>Alle salg Strøm " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                        else
                            doc.Add("<h2>Alle salg Strøm for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                        doc.AddRange(ranking.GetToppListLarge(dtTil, "Strom"));

                        if (appConfig.listerVisAccessories)
                        {
                            if (!datoPeriodeVelger || bg)
                                doc.Add("<h2>Alle salg Tilbehør " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                            else
                                doc.Add("<h2>Alle salg Tilbehør for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                            doc.AddRange(ranking.GetToppListLarge(dtTil, "Tilbehør"));
                        }

                        if (appConfig.listerVisInntjen)
                        {
                            if (!datoPeriodeVelger || bg)
                                doc.Add("<h2>Alle salg Inntjening " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h2>");
                            else
                                doc.Add("<h2>Alle salg Inntjening for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h2>");
                            doc.AddRange(ranking.GetToppListLarge(dtTil, "Inntjen"));
                        }
                    }

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingLister, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingLister);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingLister);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildVinnRankingSelger(string selgerArg)
        {
            Log.n("Henter vinnprodukt transaksjoner for selger..");
            try
            {
                webRanking.Navigate(htmlLoading);
                var doc = new List<string>();
                DateTime dtPick = pickerRankingDate.Value;
                DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;

                var ranking = new RankingVinn(this, dtFra, dtTil, dtPick);

                GetHtmlStart(doc, false, "Ranking: Vinnprodukter Selger");

                doc.Add("<h1>Vinnprodukter (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                doc.Add("<h2>Transaksjoner for selger: " + salesCodes.GetNavn(selgerArg) + "</h2>");
                doc.AddRange(ranking.GetTableHtmlSelger(selgerArg));

                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Ranking stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                }
                else
                {
                    File.WriteAllLines(htmlRankingVinnSelger, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRankingVinnSelger);
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);

                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av vinnprodukt transaksjoner for kunde.", ex);
                errorMsg.ShowDialog();

            }
        }

        private void BuildVinnRanking(bool bg = false)
        {
            string katArg = "Vinnprodukter";
            bool abort = HarSisteVersjon(katArg, appConfig.strVinnprodukter);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "]..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingVinn(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, false, "Ranking: " + katArg);

                    doc.Add("<h1>Vinnprodukter (" + avdeling.Get(appConfig.Avdeling) + ")</h1>");

                    if (ranking.tableVinn != null)
                    {
                        if (ranking.tableVinn.Rows.Count > 0)
                        {
                            if (!appConfig.vinnEnkelModus)
                            {
                                if (!datoPeriodeVelger || bg)
                                {
                                    if (appConfig.vinnDatoFraTil)
                                        if (appConfig.vinnTo > appConfig.dbTo)
                                            doc.Add("<h3>Lyd & Bilde, Tele og Data selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                        else
                                            doc.Add("<h3>Lyd & Bilde, Tele og Data selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.vinnTo.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                    else
                                        doc.Add("<h3>Lyd & Bilde, Tele og Data selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                                }
                                else
                                    doc.Add("<h3>Lyd & Bilde, Tele og Data selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                doc.AddRange(ranking.GetTableHtml(BudgetCategory.Cross));

                                if (!datoPeriodeVelger || bg)
                                {
                                    if (appConfig.vinnDatoFraTil)
                                        if (appConfig.vinnTo > appConfig.dbTo)
                                            doc.Add("<h3>Alle MDA og SDA selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                        else
                                            doc.Add("<h3>Alle MDA og SDA selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.vinnTo.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                    else
                                        doc.Add("<h3>MDA og SDA selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                                }
                                else
                                    doc.Add("<h3>MDA og SDA selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                doc.AddRange(ranking.GetTableHtml(BudgetCategory.MDASDA));

                                if (!datoPeriodeVelger || bg)
                                {
                                    if (appConfig.vinnDatoFraTil)
                                        if (appConfig.vinnTo > appConfig.dbTo)
                                            doc.Add("<h3>Kasse selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                        else
                                            doc.Add("<h3>Kasse selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.vinnTo.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                    else
                                        doc.Add("<h3>Kasse selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                                }
                                else
                                    doc.Add("<h3>Kasse selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                doc.AddRange(ranking.GetTableHtml(BudgetCategory.Kasse));

                                if (appConfig.vinnVisAftersales)
                                {
                                    if (!datoPeriodeVelger || bg)
                                    {
                                        if (appConfig.vinnDatoFraTil)
                                            if (appConfig.vinnTo > appConfig.dbTo)
                                                doc.Add("<h3>Aftersales selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                            else
                                                doc.Add("<h3>Aftersales selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.vinnTo.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                        else
                                            doc.Add("<h3>Aftersales selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                                    }
                                    else
                                        doc.Add("<h3>Aftersales selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                    doc.AddRange(ranking.GetTableHtml(BudgetCategory.Aftersales));
                                }
                            }
                            else
                            {
                                if (!datoPeriodeVelger || bg)
                                {
                                    if (appConfig.vinnDatoFraTil)
                                        if (appConfig.vinnTo > appConfig.dbTo)
                                            doc.Add("<h3>Alle selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                        else
                                            doc.Add("<h3>Alle selgere fra " + appConfig.vinnFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.vinnTo.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                    else
                                        doc.Add("<h3>Alle selgere " + dtPick.ToString("MMMM yyyy") + " (MTD " + dtPick.ToString("dddd d.", norway) + ")</h3>");
                                }
                                else
                                    doc.Add("<h3>Alle selgere for perioden fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                                doc.AddRange(ranking.GetTableHtml(BudgetCategory.None));
                            }
                        }
                        else
                            doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner eller valgt periode har ennå ikke startet.</span><br>");
                    }

                    if (appConfig.vinnVisVarekoder)
                        doc.AddRange(ranking.GetTableVarekoderHtml());

                    if (appConfig.vinnVisVarekoderExtra)
                        doc.AddRange(ranking.GetTableVarekoderEkstraHtml());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlRankingVinn, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlRankingVinn);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlRankingVinn);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        public string GetRankModeText(int rankMode)
        {
            switch (rankMode)
            {
                case 0:
                    return "Dag";
                case 1:
                    return "MTD";
                case 2:
                    return "Bonus";
                case 3:
                    return "YTD";
                case 4:
                    return "Uke";
                case 5:
                    return "År";
            }
            return "Ukjent";
        }

        private void BuildAvdTjenester(bool bg = false)
        {
            string katArg = "Tjenester";
            bool abort = HarSisteVersjon(katArg, appConfig.strTjenester);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "] ..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (appConfig.dbTo.Month != dtPick.Month && appConfig.dbTo.Year != dtPick.Year) // Hent hele måneden hvis vi IKKE er i siste måneden
                        dtTil = GetLastDayOfMonth(dtPick);
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingAvdTjenester(this, dtFra, dtTil, dtPick);

                    openXml.DeleteDocument("Tjenester", dtPick);

                    GetHtmlStart(doc, true, "Oversikt tjenester");

                    doc.Add("<h1>Oversikt tjenester</h1>");

                    string[] servicesStr = new string[4] { "Finans", "TA", "Strom", "Knowhow" };
                    foreach (string service in servicesStr)
                    {
                        doc.Add("<span class='Loading'>Beregner..</span>");
                        if (!bg && timewatch.ReadyForRefresh())
                            webRanking.DocumentText = string.Join(null, doc.ToArray());
                        doc.RemoveAt(doc.Count - 1);
                        doc.AddRange(ranking.GetTableHtmlPage(appConfig.rankingAvdelingMode, service));
                    }

                    doc.Add(Resources.htmlEnd);
                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlAvdTjenester, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlAvdTjenester);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlAvdTjenester);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }

        private void BuildAvdSnittpriser(bool bg = false)
        {
            string katArg = "Snittpriser";
            bool abort = HarSisteVersjon(katArg, appConfig.strSnittpriser);
            try
            {
                if (!bg && !abort) timewatch.Start();
                if (!bg)
                    savedPage = katArg;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + katArg + "] ..");
                    if (!bg)
                        webRanking.Navigate(htmlLoading);
                    var doc = new List<string>();
                    DateTime dtPick = pickerRankingDate.Value;
                    DateTime dtFra = GetFirstDayOfMonth(dtPick); DateTime dtTil = dtPick;
                    if (appConfig.dbTo.Month != dtPick.Month && appConfig.dbTo.Year != dtPick.Year) // Hent hele måneden hvis vi IKKE er i siste måneden
                        dtTil = GetLastDayOfMonth(dtPick);
                    if (datoPeriodeVelger && !bg)
                    {
                        dtFra = datoPeriodeFra;
                        dtTil = datoPeriodeTil;
                    }

                    var ranking = new RankingAvdSnittpriser(this, dtPick);

                    openXml.DeleteDocument(katArg, dtPick);

                    GetHtmlStart(doc, true, "Ranking: " + katArg);

                    doc.Add("<h1>Snittpriser hovedprodukter</h1>");

                    if (!datoPeriodeVelger || bg)
                        doc.Add("<h3>Periode: " + GetRankModeText(appConfig.rankingAvdelingMode) + " " + dtPick.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    else
                        doc.Add("<h3>Periode: Fra " + dtFra.ToString("dddd d. MMMM yyyy", norway) + " til " + dtTil.ToString("dddd d. MMMM yyyy", norway) + "</h3>");
                    doc.Add("<span class='Loading'>Beregner..</span>");
                    if (!bg && timewatch.ReadyForRefresh())
                        webRanking.DocumentText = string.Join(null, doc.ToArray());
                    doc.RemoveAt(doc.Count - 1);

                    doc.AddRange(ranking.GetTableHtml(appConfig.rankingAvdelingMode));

                    doc.AddRange(ranking.GetVaregrupper());

                    doc.Add(Resources.htmlEnd);

                    if (stopRanking)
                    {
                        stopRanking = false;
                        ClearHash(katArg);
                        Log.n("Ranking stoppet.", Color.Red);
                        webRanking.Navigate(htmlStopped);
                    }
                    else
                    {
                        if (datoPeriodeVelger && !bg)
                        {
                            File.WriteAllLines(htmlPeriode, doc.ToArray(), Encoding.Unicode);
                            webRanking.Navigate(htmlPeriode);
                        }
                        else
                        {
                            File.WriteAllLines(htmlAvdSnittpriser, doc.ToArray(), Encoding.Unicode);
                            if (!bg)
                                webRanking.Navigate(htmlAvdSnittpriser);
                            if (!bg) Log.n("Ranking [" + katArg + "] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                        }
                    }
                }
                else if (!bg)
                    webRanking.Navigate(htmlAvdSnittpriser);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!bg)
                {
                    webRanking.Navigate(htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
        }


        private void MakeReport(string katArg = "")
        {
            if (!String.IsNullOrEmpty(katArg))
                RunReport(katArg);
            else
            {
                var page = currentPage();
                if (page == "Butikk" || page == "Data" || page == "Tele" || page == "AudioVideo" || page == "KnowHow")
                    RunReport(page);
                else if (page == "Rapport")
                    Log.n("Velg en kategori først.", Color.Black);
                else
                    Log.n("Beklager, kan ikke lage rapport av gjeldene side.", Color.Black);
            }
        }

        private void RunReport(string katArg)
        {
            if (!bwRanking.IsBusy && !bwReport.IsBusy && (katArg == "Butikk" || katArg == "Data" || katArg == "AudioVideo" || katArg == "Tele" || katArg == "KnowHow"))
            {
                if (!EmptyDatabase())
                {
                    groupRankingChoices.Enabled = false;
                    Log.n("Lager [" + katArg + "] rapport ..", Color.Black, false, true);
                    bwReport.RunWorkerAsync(katArg);
                    HighlightButton();
                }
                else
                    webRanking.Navigate(htmlImport);
            }
        }

        private void bwReport_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string katArg = (string)e.Argument;
            if (katArg == "Butikk")
                makeButikkReport();
            else if (katArg == "KnowHow")
                makeKnowHowReport();
            else if (katArg == "Data")
                makeDataReport();
            else if (katArg == "AudioVideo")
                makeAudioVideoReport();
            else if (katArg == "Tele")
                makeTeleReport();
            else
                Log.n("Ingen kategori valgt for rapport generering.");
        }

        private void bwReport_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
                Log.Status("Klar.");
            groupRankingChoices.Enabled = true;
        }

        private void makeKnowHowReport()
        {
            try
            {
                timewatch.Start();
                webRanking.Navigate(htmlLoading);
                DateTime dtPick = pickerRankingDate.Value;
                KnowHowReport report = new KnowHowReport(this, appConfig.Avdeling);
                var doc = new List<string>();
                GetHtmlStart(doc, true);

                doc.Add("<span style='font-size:13.0pt;font-weight:400;font-style:bold;text-decoration:none;font-family:Calibri, sans-serif;float:left'>Avdeling: KnowHow (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                doc.Add("<span class='xTitle' style='font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;'>" +
                        "Rapport for periode fra " + appConfig.dbFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.dbTo.ToString("dddd d. MMMM yyyy", norway) + "</span><br>");

                report.makeMonthly(dtPick, bwReport, doc);
                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Rapport stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                    timewatch.Stop();
                }
                else
                {
                    File.WriteAllLines(htmlRapport, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRapport);
                    Log.n("Rapport [KnowHow] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch (Exception ex)
            {
                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av rapport for [KnowHow]", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void makeDataReport()
        {
            try
            {
                timewatch.Start();
                webRanking.Navigate(htmlLoading);
                DateTime dtPick = pickerRankingDate.Value;
                DataReport report = new DataReport(this, appConfig.Avdeling);
                var doc = new List<string>();
                GetHtmlStart(doc, true);

                doc.Add("<span style='font-size:13.0pt;font-weight:400;font-style:bold;text-decoration:none;font-family:Calibri, sans-serif;float:left'>Avdeling: Data (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                doc.Add("<span class='xTitle' style='font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;'>" +
                        "Rapport for periode fra " + appConfig.dbFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.dbTo.ToString("dddd d. MMMM yyyy", norway) + "</span><br>");

                report.makeMonthly(dtPick, bwReport, doc);
                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Rapport stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                    timewatch.Stop();
                }
                else
                {
                    File.WriteAllLines(htmlRapport, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRapport);
                    Log.n("Rapport [Data] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch (Exception ex)
            {
                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av rapport for [Data]", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void makeAudioVideoReport()
        {
            try
            {
                timewatch.Start();
                webRanking.Navigate(htmlLoading);
                DateTime dtPick = pickerRankingDate.Value;
                AudioVideoReport report = new AudioVideoReport(this, appConfig.Avdeling);
                var doc = new List<string>();
                GetHtmlStart(doc, true);

                doc.Add("<span style='font-size:13.0pt;font-weight:400;font-style:bold;text-decoration:none;font-family:Calibri, sans-serif;float:left'>Avdeling: Lyd og Bilde (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                doc.Add("<span class='xTitle' style='font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;'>" +
                        "Rapport for periode fra " + appConfig.dbFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.dbTo.ToString("dddd d. MMMM yyyy", norway) + "</span><br>");
                
                report.makeMonthly(dtPick, bwReport, doc);
                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Rapport stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                    timewatch.Stop();
                }
                else
                {
                    File.WriteAllLines(htmlRapport, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRapport);
                    Log.n("Rapport [AudioVideo] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch (Exception ex)
            {
                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av rapport for [AudioVideo]", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void makeTeleReport()
        {
            try
            {
                timewatch.Start();
                webRanking.Navigate(htmlLoading);
                DateTime dtPick = pickerRankingDate.Value;
                TeleReport report = new TeleReport(this, appConfig.Avdeling);
                var doc = new List<string>();
                GetHtmlStart(doc, true);

                doc.Add("<span style='font-size:13.0pt;font-weight:400;font-style:bold;text-decoration:none;font-family:Calibri, sans-serif;float:left'>Avdeling: Tele (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                doc.Add("<span class='xTitle' style='font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;'>" +
                        "Rapport for periode fra " + appConfig.dbFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.dbTo.ToString("dddd d. MMMM yyyy", norway) + "</span><br>");
                
                report.makeMonthly(dtPick, bwReport, doc);
                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Rapport stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                    timewatch.Stop();
                }
                else
                {
                    File.WriteAllLines(htmlRapport, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRapport);
                    Log.n("Rapport [Tele] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch (Exception ex)
            {
                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av rapport for [Tele]", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void makeButikkReport()
        {
            try
            {
                timewatch.Start();
                webRanking.Navigate(htmlLoading);
                DateTime dtPick = pickerRankingDate.Value;
                ButikkReport report = new ButikkReport(this, appConfig.Avdeling);
                var doc = new List<string>();
                GetHtmlStart(doc, true);
                doc.Add("<span style='font-size:13.0pt;font-weight:400;font-style:bold;text-decoration:none;font-family:Calibri, sans-serif;float:left'>Oversikt butikk (" + avdeling.Get(appConfig.Avdeling) + ")</span><span style='font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;color:gray;float:right'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br><br>");
                doc.Add("<span class='xTitle' style='font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;'>" +
                        "Rapport for periode fra " + appConfig.dbFrom.ToString("dddd d. MMMM yyyy", norway) + " til " + appConfig.dbTo.ToString("dddd d. MMMM yyyy", norway) + "</span><br>");
                
                report.makeMonthly(dtPick, bwReport, doc);
                doc.Add(Resources.htmlEnd);

                if (stopRanking)
                {
                    stopRanking = false;
                    Log.n("Rapport stoppet.", Color.Red);
                    webRanking.Navigate(htmlStopped);
                    timewatch.Stop();
                }
                else
                {
                    File.WriteAllLines(htmlRapport, doc.ToArray(), Encoding.Unicode);
                    webRanking.Navigate(htmlRapport);
                    Log.n("Rapport [Butikk] tok " + timewatch.Stop() + " sekunder.", Color.Black, true);
                }
            }
            catch(Exception ex)
            {
                webRanking.Navigate(htmlError);
                FormError errorMsg = new FormError("Feil ved generering av rapport for [Butikk]", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private string StringRankingDato(DateTime date)
        {
            return date.ToString("MMMM yyyy") + " (MTD " + date.ToString("dddd d.", norway) + ")";
        }
    }
}

