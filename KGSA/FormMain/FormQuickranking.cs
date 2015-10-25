using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    partial class FormMain
    {
        public void startDelayedDailyImport()
        {
            if (IsBusy())
                return;

            worker = new BackgroundWorker();
            worker.DoWork += new DoWorkEventHandler(bwDaily_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwDaily_Completed);

            processing.SetVisible = true;
            processing.SetProgressStyle = ProgressBarStyle.Continuous;
            processing.SetBackgroundWorker = worker;
            for (int b = 0; b < 100; b++)
            {
                processing.SetText = "Starter daglig budsjett import om " + (((b / 10) * -1) + 10) + " sekunder..";
                processing.SetValue = b;
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);

                if (processing.userPushedCancelButton)
                {
                    Log.n("Brukeren avbrøt handlingen.");
                    return;
                }
            }

            RestoreWindow();

            processing.SetProgressStyle = ProgressBarStyle.Marquee;
            processing.SetBackgroundWorker = worker;
            processing.SetText = "Kjører daglig budsjett rutine..";
            Log.n("Kjører daglig budsjett rutine..");
            worker.RunWorkerAsync();
        }

        private void bwDaily_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;

            try
            {
                FormMacro formMacro = (FormMacro)StartMacro(appConfig.dbTo, macroProgramQuick, bwQuickAuto, 0);
                if (formMacro == null || formMacro.errorCode != 0)
                {
                    if (tableMacroQuick == null || tableMacroQuick.Rows.Count < 5)
                    {
                        Log.n("Daglig Budsjett: Feil oppstod under makro kjøring. Mangler data fra Elguide. Se logg for detaljer", Color.Red);
                        e.Result = 7;
                        return;
                    }
                    else
                    {
                        Log.n("Daglig Budsjett: Feil oppstod under makro kjøring. Feilbeskjed: " + formMacro.errorMessage + " Kode: " + formMacro.errorCode, Color.Red);
                        e.Result = formMacro.errorCode;
                        return;
                    }
                }

                MakeBudgetPage(BudgetCategory.Daglig, worker);

                e.Result = 0;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Daglig Budsjett: Feil oppstod under konvertering og sending av kveldsranking til PDF. Se logg for detaljer.", Color.Red);
            }
            e.Result = 8;
        }

        private void bwDaily_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            autoMode = false;
            ProgressStop();

            if (e.Error != null)
            {
                Log.n("Kveldstall: Ukjent feil oppstod. " + e.Error.Message + " Se logg for detaljer.", Color.Red);
            }
            else if (e.Cancelled || (int)e.Result == 2)
            {
                Log.n("Kveldstall: Avbrutt av bruker.");
            }
            else if ((int)e.Result != 0)
            {
                Log.n("Kveldstall: Ukjent feil oppstod. Se logg for detaljer.", Color.Red);
            }
            else
            {
                Log.n("Kveldstall: Fullført uten feil.", Color.Green);
                processing.SetVisible = false;
                return;
            }
            processing.HideDelayed();
            this.Activate();
        }

        public void startDelayedAutoQuickImport()
        {
            if (IsBusy())
                return;

            processing.SetVisible = true;
            processing.SetProgressStyle = ProgressBarStyle.Continuous;
            processing.SetBackgroundWorker = bwQuickAuto;
            for (int b = 0; b < 100; b++)
            {
                processing.SetText = "Starter automatisk kvelds ranking om " + (((b / 10) * -1) + 10) + " sekunder..";
                processing.SetValue = b;
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);

                if (processing.userPushedCancelButton)
                {
                    Log.n("Brukeren avbrøt handlingen.");
                    return;
                }
            }

            RestoreWindow();

            processing.SetProgressStyle = ProgressBarStyle.Marquee;
            processing.SetBackgroundWorker = bwQuickAuto;
            processing.SetText = "Kjører kveldsranking rutine..";
            Log.n("Kjører kveldsranking rutine..");
            bwQuickAuto.RunWorkerAsync();
        }

        private void bwQuickAuto_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;

            var macroAttempt = 0;
            var macroMaxAttempts = 4;

            retrymacro:
            macroAttempt++;

            FormMacro formMacro = (FormMacro)StartMacro(appConfig.dbTo, macroProgramQuick, bwQuickAuto, macroAttempt);
            e.Result = formMacro.errorCode;
            if (formMacro.errorCode == 6)
            {
                Log.n("En kritisk feil forhindret makro i å utføre sine oppgaver. Se logg for detaljer.", Color.Red);
                e.Result = 6;
                return;
            }

            if (tableMacroQuick == null || tableMacroQuick.Rows.Count < 5 || formMacro.errorCode != 0)
            {
                System.Threading.Thread.Sleep(3000);
                if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                    goto retrymacro;
                if (formMacro.errorCode != 2)
                    Log.n("Kveldstall: Feil oppstod under kveldsranking automatisering. Feilbeskjed: " + formMacro.errorMessage + " Kode: " + formMacro.errorCode, Color.Red);
                return;
            }

            BudgetImporter budgetImporter = new BudgetImporter(this, DateTime.Now);

            if (appConfig.dailyBudgetQuickRankingAutoUpdate && appConfig.dailyBudgetIncludeInQuickRanking)
            {
                Log.n("Kveldstall: Starter nedlasting av dagens budsjett..");
                processing.SetText = "Henter dagens budsjett..";
                budgetImporter.FindAndDownloadBudget(bwQuickAuto);
            }

            ClearBudgetHash(BudgetCategory.Daglig);
            MakeBudgetPage(BudgetCategory.Daglig, bwQuickAuto);

            Log.n("Kveldstall: Konverterer kveldranking til PDF..");
            processing.SetText = "Konverterer kveldsranking til PDF..";

            string pdfBudget = CreatePDF("", "\"" + settingsPath + "\\budsjettDaglig.html\" ", bwQuickAuto);
            if (String.IsNullOrEmpty(pdfBudget) || !File.Exists(pdfBudget))
            {
                // Feil under pdf generering
                Log.n("Kveldstall: Feil oppstod under konvertering av kveldsranking til PDF. Se logg for detaljer.", Color.Red);
                e.Result = 6;
                return;
            }

            processing.SetText = "Sender kveldstall..";
            KgsaEmail email = new KgsaEmail(this);
            if (!email.Send(pdfBudget, DateTime.Now, "Quick", appConfig.epostEmneQuick, appConfig.epostBodyQuick))
            {
                // Feil under pdf generering
                Log.n("Kveldstall: Feil oppstod under sending av kveldsranking til PDF. Se logg for detaljer.", Color.Red);
                e.Result = 8;
                return;
            }
        }

        private void bwQuickAuto_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;
            if (e.Error != null)
            {
                Log.n("Kveldstall: Ukjent feil oppstod. " + e.Error.Message + " Se logg for detaljer.", Color.Red);
            }
            else if (e.Cancelled || (e.Result != null && (int)e.Result == 2))
            {
                Log.n("Kveldstall: Avbrutt av bruker.");
            }
            else if (e.Result != null && (int)e.Result != 0)
            {
                Log.n("Kveldstall: Ukjent feil oppstod. Se logg for detaljer.", Color.Red);
            }
            else
            {
                Log.n("Kveldstall: Fullført uten feil.", Color.Green);
            }
            processing.HideDelayed();
            this.Activate();
        }
    }
}
