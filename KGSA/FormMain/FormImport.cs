﻿using System;
using System.Data;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using FileHelpers;
using KGSA.Properties;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Drawing;
using System.ComponentModel;

namespace KGSA
{
    partial class FormMain
    {
        private void bwMacroRanking_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;
            var macroAttempt = 1;
            var macroMaxAttempts = 4;

        retrymacro:

            Logg.Log("Auto: Kjører makro.. [" + macroAttempt + "]");

            DateTime date = dbTilDT;
            double span = (DateTime.Now - dbTilDT).TotalDays;
            if (span > 31)
                date = DateTime.Now.AddMonths(-1); // Begrens oss til å importere en måned bak i tid
            if (dbTilDT == DateTime.Now)
                date = GetFirstDayOfMonth(dbTilDT);

            var macroForm = (FormMacro)StartMacro(date, macroProgram, bwMacroRanking, macroAttempt);
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

                e.Result = macroForm.errorCode;
                return;
            }

            if (!File.Exists(appConfig.csvElguideExportFolder + @"\irank.csv") || File.GetLastWriteTime(appConfig.csvElguideExportFolder + @"\irank.csv").Date != DateTime.Now.Date)
            {
                // CSV finnes ikke eller filen er ikke oppdatert i dag i.e. data ble ikke eksportert riktig med makro
                macroAttempt++;
                Logg.Log("Auto: CSV er IKKE oppdatert, eller ingen tilgang. Sjekk CSV lokasjon og makro innstillinger.", Color.Red);
                for (int i = 0; i < 60; i++)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
                if (macroAttempt < macroMaxAttempts && macroForm.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                    goto retrymacro;

                return;
            }

            Logg.Log("Auto: Importerer data..");

            csvFilesToImport.Clear();
            ImportManager importMng = new ImportManager(this, csvFilesToImport);
            importMng.DoImportTransactions(bwAutoRanking, true);
            if (importMng.returnCode != 0)
            {
                // Feil under importering
                macroAttempt++;
                Logg.Log("Auto: Feil oppstod under importering. Kode: " + importMng.returnCode, Color.Red);
                for (int i = 0; i < 60; i++)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
                if (macroAttempt < macroMaxAttempts && macroForm.errorCode != 2)
                    goto retrymacro;

                return;
            }
        }

        private void bwMacroRanking_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;
            if (e.Cancelled)
            {
                Logg.Log("Makro: Oppgave kansellert.");
            }
            else if (e.Error != null)
            {
                Logg.Log("Makro: Oppgave avsluttet med feil " + e.Error.Message, Color.Red);
            }
            else
            {
                int returnCode = -1;
                if (e.Result != null)
                    returnCode = (int)e.Result;
                else
                    returnCode = 0;

                if (returnCode == 0)
                {
                    RetrieveDb();
                    Reload(true);
                    UpdateUi();
                    Logg.Log("Makro: Oppgave fullført.", Color.Green);
                }
                else if (returnCode == 2)
                {
                    Logg.Log("Makro: Oppgave avbrutt av bruker.");
                }
                else if (returnCode == -1)
                {
                    Logg.Log("Makro: Oppgave avsluttet med ukjent status.", Color.OrangeRed);
                }
                else
                {
                    Logg.Debug("Makro: Ukjent returkode: " + returnCode);
                }
            }
            processing.HideDelayed();
            this.Activate();
        }

        private void velgObsoleteCSV()
        {
            // Browse etter obsolete.csv
            try
            {
                var fdlg = new OpenFileDialog();
                fdlg.Title = "Velg Obsolete CVS-fil eksportert fra Elguide";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "Alle filer (*.*)|*.*|ZIP filer (*.zip)|*.zip|CVS filer (*.csv)|*.csv";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                fdlg.Multiselect = false;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    if (fdlg.FileName.EndsWith(".zip"))
                    {
                        processing.SetVisible = true;
                        ProgressStart();
                        processing.SetText = "Utpakking..";
                        string extracted = obsolete.Decompress(fdlg.FileName);
                        if (extracted != "")
                            RunObsoleteImport(extracted);
                        else
                        {
                            processing.SetVisible = false;
                            MessageBox.Show("Obs! Utpakking av arkiv mislykkes. Se logg for detaljer.\nPrøv igjen eller pakk ut manuelt før importering.", "KGSA - Info", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                        RunObsoleteImport(fdlg.FileName);
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                Logg.Log("Unntak ved velging av CSV fil. " + ex.Message, Color.Red);
            }
        }

        private void RunObsoleteImport(string str = "")
        {
            if (!IsBusy(true))
            {
                processing.SetVisible = true;
                processing.SetText = "Starter importering av varer..";
                processing.SetValue = 25;
                processing.SetBackgroundWorker = bwImportObsolete;
                bwImportObsolete.RunWorkerAsync(str);
                timewatch.Start();
            }
        }

        private void bwImportObsolete_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string str = (string)e.Argument;
            bool complete = false;

            if (str != "" && str != null)
                complete = obsolete.Import(str, bwImportObsolete);
            else
                complete = obsolete.Import(appConfig.csvElguideExportFolder + @"\obsolete.csv", bwImportObsolete);

            e.Result = complete;
        }

        public void bwProgressReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressReport(Convert.ToDecimal(e.ProgressPercentage), (StatusProgress)e.UserState);
        }

        private void bwImportObsolete_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            try
            {
                bool complete = (bool)e.Result;
                if (complete)
                {
                    var tid = timewatch.Stop();
                    Logg.Log("Importering tok " + tid + " sekunder.", Color.Black, true);
                    database.ClearCacheTables();
                    RetrieveDbStore();
                    ReloadStore(true);
                    UpdateUi();
                    Logg.Log("Vare import: Importering fullført uten feil.", Color.Green);
                    processing.SetText = "Fullført!";
                }
                else
                {
                    Logg.Log("Vare import: Importering fullført med feil! Se logg for detaljer.", Color.Red);
                    processing.SetText = "Avbrutt..";
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            processing.HideDelayed();
            this.Activate();
        }

        private void RunImport(bool ignoreBusy = false)
        {
            if (!IsBusy(true) || ignoreBusy)
            {
                processing.SetVisible = true;
                processing.SetBackgroundWorker = bwImport;
                bwImport.RunWorkerAsync();
                processing.SetText = "Starter importering..";
                timewatch.Start();
            }
        }

        private void bwImport_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();

            ImportManager importMng = new ImportManager(this, csvFilesToImport);
            importMng.DoImportTransactions(bwImport, false);

            e.Result = importMng;
            csvFilesToImport.Clear();
        }

        private void bwImport_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            try
            {
                var tid = timewatch.Stop();
                Logg.Log("Importering tok " + tid + " sekunder.", Color.Black, true);

                ImportManager importMng = (ImportManager)e.Result;

                if (importMng.returnCode == 0)
                {
                    if (importMng.importReadErrors > 0)
                        Logg.Log("Importering fullført! Transaksjoner med lesefeil: " + importMng.importReadErrors, Color.Orange);
                    else
                        Logg.Log("Importering fullført uten feil!", Color.Green);

                    RetrieveDb(true); // forced update
                    RetrieveDbStore();
                    processing.SetValue = 95;
                    UpdateUi();
                    graphPanelTop.Invalidate();
                    processing.HideDelayed();
                    this.Activate();
                    if (!autoMode)
                    {
                        ClearHash();
                        ClearBudgetHash(BudgetCategory.None);
                        string errorsStr = "";
                        if (importMng.importReadErrors > 0)
                            errorsStr = "\nLese feil: " + importMng.importReadErrors + "\nSe logg for detaljer.\n";

                        if (MessageBox.Show("Importering fullført!\n" + errorsStr + "\n" + importMng.importCount.ToString("#,##0") + " transaksjoner tok " + tid + " sekunder.\nSiste dag: " + dbTilDT.ToShortDateString() + "\n\nVil du oppdatere gjeldene ranking?", "KGSA - Importering", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                            Reload(true); // forced update
                    }
                }
                else if (importMng.returnCode == 4)
                {
                    Logg.Log("Ingen tilgang til CSV fil eller CSV inneholdte ingen transaksjoner.", Color.Red);
                    processing.SetText = "Importering feilet!";
                }
                else if (importMng.returnCode == 2)
                {
                    Logg.Log("Importering avbrutt av bruker.", Color.Red);
                    processing.SetText = "Importering avbrutt av bruker!";
                }
                else
                {
                    Logg.Log("Feil oppstod under importering av CSV", Color.Red);
                    processing.SetText = "Importering avbrutt!";
                }
                processing.HideDelayed();
                this.Activate();
            }
            catch(Exception ex)
            {
                processing.SetText = "Importering avsluttet med feil!";
                processing.HideDelayed();
                Logg.Unhandled(ex);
            }
        }

        /// <summary>
        /// Starter macro form
        /// </summary>
        /// <param name="dateArg">Dato på hva som skal importeres</param>
        /// <param name="macroProgramArg">Makro program som skal benyttes</param>
        /// <param name="attemptsArg">Hvilket forsøkt dette er</param>
        /// <returns>Form som ble brukt for å hente ut returkoder</returns>
        public Form StartMacro(DateTime dateArg, string macroProgramArg, BackgroundWorker bw, int attemptsArg = 0, bool ignoreExtraWait = false)
        {
            Form form = new FormMacro(appConfig, dateArg, macroProgramArg, attemptsArg, ignoreExtraWait, bw);
            form.StartPosition = FormStartPosition.CenterScreen;
            form.ShowDialog();

            return form;
        }

        private void delayedMacroRankingImport()
        {
            try
            {
                processing.SetVisible = true;
                processing.SetProgressStyle = ProgressBarStyle.Continuous;
                processing.SetBackgroundWorker = bwMacroRanking;
                for (int b = 0; b < 100; b++)
                {
                    processing.SetText = "Starter automatisk ranking import om " + (((b / 10) * -1) + 10) + " sekunder..";

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

                RunAutoMacroRankingImport();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void RunAutoMacroRankingImport()
        {
            if (!IsBusy(true))
            {
                processing.SetVisible = true;
                processing.SetText = "Importerer transaksjoner med makro..";
                Logg.Log("Importerer transaksjoner med makro..");
                processing.SetValue = 25;
                processing.SetBackgroundWorker = bwMacroRanking;
                bwMacroRanking.RunWorkerAsync();
            }
        }
    }
}