using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using KGSA.Properties;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KGSA
{
    partial class FormMain
    {
        public void ShowHideGui_Macro(bool show)
        {
            if (show)
            {
                // from buttons
                buttonRankingMakro.Enabled = true;
                buttonBudgetActionMacroImport.Enabled = true;
                buttonLagerMakro.Enabled = true;
                buttonServiceMacro.Enabled = true;

                // from menu
                hentTransaksjonerToolStripMenuItem.Enabled = true;
                hentLagervarerToolStripMenuItem.Enabled = true;
                hentServicerToolStripMenuItem.Enabled = true;
                kjørAutorankingToolStripMenuItem.Enabled = true;
                kjørKveldstankingToolStripMenuItem.Enabled = true;
            }
            else
            {
                // from buttons
                buttonRankingMakro.Enabled = false;
                buttonBudgetActionMacroImport.Enabled = false;
                buttonLagerMakro.Enabled = false;
                buttonServiceMacro.Enabled = false;

                // from menu
                hentTransaksjonerToolStripMenuItem.Enabled = false;
                hentLagervarerToolStripMenuItem.Enabled = false;
                hentServicerToolStripMenuItem.Enabled = false;
                kjørAutorankingToolStripMenuItem.Enabled = false;
                kjørKveldstankingToolStripMenuItem.Enabled = false;
            }
        }

        public void ShowHideGui_FullTrans(bool show)
        {
            if (show)
            {
                buttonOversikt.Enabled = true;
                buttonLister.Enabled = true;
                buttonRankVinnprodukter.Enabled = true;
                buttonAvdTjenester.Enabled = true;
                buttonAvdSnittpriser.Enabled = true;
            }
            else
            {
                buttonOversikt.Enabled = false;
                buttonLister.Enabled = false;
                buttonRankVinnprodukter.Enabled = false;
                buttonAvdTjenester.Enabled = false;
                buttonAvdSnittpriser.Enabled = false;
            }
        }

        public void ShowHideGui_EmptyRanking(bool show)
        {
            if (show)
            {
                groupRankingPages.Enabled = true;
                groupRankingChoices.Enabled = true;
                buttonOpenExcel.Enabled = true;
                buttonRankingActionOpenPdf.Enabled = true;

                groupBudgetPages.Enabled = true;
                groupBudgetChoices.Enabled = true;
                buttonBudgetActionOpenExcel.Enabled = true;
                buttonBudgetActionOpenPdf.Enabled = true;

                groupGraphPages.Enabled = true;
                groupGraphChoices.Enabled = true;
                buttonGraphCopyToClipboard.Enabled = true;


                toolStripMenuItemSaveFullPdf.Enabled = true;

                butikkToolStripMenuItem.Enabled = true;
                dataToolStripMenuItem.Enabled = true;
                lydOgBildeToolStripMenuItem.Enabled = true;
                teleToolStripMenuItem.Enabled = true;

                panelDatabaseSearch.Enabled = true;
            }
            else
            {
                groupRankingPages.Enabled = false;
                groupRankingChoices.Enabled = false;
                buttonOpenExcel.Enabled = false;
                buttonRankingActionOpenPdf.Enabled = false;

                groupBudgetPages.Enabled = false;
                groupBudgetChoices.Enabled = false;
                buttonBudgetActionOpenExcel.Enabled = false;
                buttonBudgetActionOpenPdf.Enabled = false;

                groupGraphPages.Enabled = false;
                groupGraphChoices.Enabled = false;
                buttonGraphCopyToClipboard.Enabled = false;

                toolStripMenuItemSaveFullPdf.Enabled = false;

                butikkToolStripMenuItem.Enabled = false;
                dataToolStripMenuItem.Enabled = false;
                lydOgBildeToolStripMenuItem.Enabled = false;
                teleToolStripMenuItem.Enabled = false;

                panelDatabaseSearch.Enabled = false;
            }
        }

        public void ShowHideGui_EmptyStore(bool show)
        {
            if (show)
            {
                groupBoxLagerKat.Enabled = true;
                groupBoxLagerValg.Enabled = true;

                buttonLagerExcel.Enabled = true;
                buttonStoreOpenPdf.Enabled = true;
                buttonLagerOppPrisguide.Enabled = true;
                buttonLagerOppUkeannonser.Enabled = true;
            }
            else
            {
                groupBoxLagerKat.Enabled = false;
                groupBoxLagerValg.Enabled = false;

                buttonLagerExcel.Enabled = false;
                buttonStoreOpenPdf.Enabled = false;
                buttonLagerOppPrisguide.Enabled = false;
                buttonLagerOppUkeannonser.Enabled = false;
            }
        }

        public void ShowHideGui_EmptyService(bool show)
        {
            if (show)
            {
                groupServicePages.Enabled = true;
                groupServiceOptions.Enabled = true;

                buttonServiceActionOpenPdf.Enabled = true;

                nullstillBehandletMarkeringToolStripMenuItem.Enabled = true;
            }
            else
            {
                groupServicePages.Enabled = false;
                groupServiceOptions.Enabled = false;

                buttonServiceActionOpenPdf.Enabled = false;

                nullstillBehandletMarkeringToolStripMenuItem.Enabled = false;
            }
        }

        #region WEB STUFF
        private void StartWebserver()
        {
            try
            {
                if (appConfig.webserverEnabled && appConfig.webserverPort >= 1024 && appConfig.webserverPort <= 65535 && !String.IsNullOrEmpty(appConfig.webserverHost))
                {
                    server = new KgsaServer(this);
                    server.StartWebserver();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        delegate void SetWebStartProcessCallback(string command);
        public void WebStartProcess(string value)
        {
            if (this.InvokeRequired)
            {
                SetWebStartProcessCallback d = new SetWebStartProcessCallback(WebStartProcess);
                this.Invoke(d, new object[] { value });
            }
            else
            {
                if (value == "importranking")
                {
                    if (!IsBusy())
                        delayedMacroRankingImport();
                }
                else if (value == "importstore")
                {
                    if (!IsBusy())
                        delayedAutoStore();
                }
                else if (value == "importservice")
                {
                    if (!IsBusy())
                        delayedAutoServiceImport();
                }
                else if (value == "update")
                {
                    if (!IsBusy())
                        RunCreateHtml();
                }
                else if (value == "autoranking")
                {
                    if (!IsBusy())
                        delayedAutoRanking();
                }
            }
        }

        #endregion


        private void OpenSendEmail()
        {
            if (String.IsNullOrEmpty(appConfig.epostAvsender) || String.IsNullOrEmpty(appConfig.epostSMTPserver))
            {
                Logg.Log("Mangler epost opplysninger. Sjekk innstillinger!", Color.Red);
                return;
            }

            FormSendEmail se = new FormSendEmail(this);

            if (se.ShowDialog() == DialogResult.OK)
            {
                Logg.Log("Starter sending av ranking..");

                processing.SetVisible = true;
                processing.SetBackgroundWorker = bwSendEmail;
                bwSendEmail.RunWorkerAsync(se);
            }
        }

        private void bwSendEmail_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            FormSendEmail se = (FormSendEmail)e.Argument;
            KgsaEmail kgsaEmail = new KgsaEmail(this);

            string rankingType = "";
            if (se.comboBoxType.SelectedIndex == 0)
                rankingType = "Full";
            else if (se.comboBoxType.SelectedIndex == 1)
                rankingType = "AudioVideo";
            else if (se.comboBoxType.SelectedIndex == 2)
                rankingType = "Telecom";
            else if (se.comboBoxType.SelectedIndex == 3)
                rankingType = "Computer";
            else if (se.comboBoxType.SelectedIndex == 4)
                rankingType = "Cross";
            else if (se.comboBoxType.SelectedIndex == 5)
                rankingType = "Vinnprodukter";
            else if (se.comboBoxType.SelectedIndex == 6)
                rankingType = "FullUkestart";

            string epostGruppe = se.comboBoxGruppe.GetItemText(se.comboBoxGruppe.SelectedItem);
            if (se.comboBoxGruppe.SelectedIndex == 0)
                epostGruppe = "";

            string pdf = CreatePDF(rankingType, "", bwAutoRanking);
            if (!String.IsNullOrEmpty(pdf))
            {
                processing.SetText = "Sender e-post for gruppe \"" + epostGruppe + "\"..";
                if (!kgsaEmail.Send(pdf, appConfig.dbTo, epostGruppe, se.textBoxTitle.Text, se.textBoxContent.Text))
                    e.Cancel = true;
            }
            else
                e.Cancel = true;
        }

        private void bwSendEmail_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            processing.HideDelayed();
            if (e.Cancelled)
                Logg.Log("Sending av e-post avbryt av bruker.");
            else if (e.Error != null)
                Logg.Log("Feil oppstod under sending av e-post. Sjekk logg.", Color.Red);

            this.Activate();
        }

        public static DateTime RetrieveLinkerTimestamp()
        {
            string filePath = System.Reflection.Assembly.GetCallingAssembly().Location;
            const int c_PeHeaderOffset = 60;
            const int c_LinkerTimestampOffset = 8;
            byte[] b = new byte[2048];
            Stream s = null;
            try
            {
                s = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally
            {
                if (s != null)
                    s.Close();
            }
            int i = BitConverter.ToInt32(b, c_PeHeaderOffset);
            int secondsSince1970 = BitConverter.ToInt32(b, i + c_LinkerTimestampOffset);
            DateTime dt = new DateTime(1970, 1, 1, 0, 0, 0);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours);
            return dt;
        }


        private void OpenBudgetSettings()
        {
            var formBudget = new FormBudgetCreation(this);
            if (formBudget.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                ClearBudgetHash(BudgetCategory.None);
                HighlightBudgetButton(BudgetCategory.None);
            }
        }

        private void bwCreateHtml_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            Logg.Log("Starter bakgrunnsjobb..", null, true);
            e.Result = CreateHtml(bwCreateHtml);
        }

        private bool CreateHtml(BackgroundWorker bw)
        {
            try
            {
                string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                if (appConfig.importSetting.StartsWith("Full"))
                {
                    BuildToppselgereRanking(true);
                    appConfig.strToppselgere = newHash;

                    if (stopRanking || IsBusy(false, true) || autoMode)
                        return false;

                    BuildOversiktRanking(true, bw);
                    appConfig.strOversikt = newHash;

                    if (stopRanking || IsBusy(false, true) || autoMode)
                        return false;
                }

                BuildButikkRanking(true, bw);
                appConfig.strButikk = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildKnowHowRanking(true);
                appConfig.strKnowHow = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildDataRanking(true);
                appConfig.strData = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildAudioVideoRanking(true, bw);
                appConfig.strAudioVideo = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildTeleRanking(true, bw);
                appConfig.strTele = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildListerRanking(true);
                appConfig.strLister = newHash;

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                BuildVinnRanking(true);
                appConfig.strVinnprodukter = newHash;

                if (!EmptyStoreDatabase())
                {
                    string newHashStore = appConfig.Avdeling + pickerLagerDato.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();

                    if (stopRanking || IsBusy(false, true) || autoMode)
                        return false;

                    BuildStoreStatus(true);
                    appConfig.strObsolete = newHashStore;

                    if (stopRanking || IsBusy(false, true) || autoMode)
                        return false;

                    BuildStoreObsoleteList(true);
                    appConfig.strObsoleteList = newHashStore;

                    if (stopRanking || IsBusy(false, true) || autoMode)
                        return false;

                    BuildStoreObsoleteImports(true);
                    appConfig.strObsoleteImports = newHashStore;
                }

                if (stopRanking || IsBusy(false, true) || autoMode)
                    return false;

                if (service != null)
                    if (service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
                        MakeServiceOversikt(true, bw);

                return true;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        private void bwCreateHtml_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (e.Error != null || e.Cancelled || !(bool)e.Result)
                Logg.Log("Bakgrunnsjobb avbrutt.");
            else
                Logg.Log("Bakgrunnsjobb ferdig.", null, true);
        }


        public void GetHtmlStart(List<string> doc, bool watermark = false, string title = "Untitled")
        {
            doc.Add("<html>");
            doc.Add("<head>");
            doc.Add("<meta charset=\"UTF-8\">");
            doc.Add("<title>KGSA - " + title + "</title>");

            doc.Add("<style id=\"stylesheet\">");
            doc.Add("body {");
	        doc.Add("    font-weight:400;");
	        doc.Add("    font-family:Calibri, sans-serif;");
	        doc.Add("    font-style:normal;");
	        doc.Add("    text-decoration:none;");
	        doc.Add("    display: block;");
	        doc.Add("    padding: 0;");
	        doc.Add("    margin: 10px;");
            if (watermark)
            {
	            doc.Add("    background-repeat:no-repeat;");
	            doc.Add("    background-attachment:fixed;");
                doc.Add("    background-position: top right;");
                if (appConfig.chainElkjop)
                    doc.Add("    background-image:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUAAAAIRCAMAAAAMdJRXAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6MkVCMEIxQTI5MTIzMTFFMzlGRDVCRERCMDgyOUMyRTIiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6MkVCMEIxQTM5MTIzMTFFMzlGRDVCRERCMDgyOUMyRTIiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDoyRUIwQjFBMDkxMjMxMUUzOUZENUJEREIwODI5QzJFMiIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDoyRUIwQjFBMTkxMjMxMUUzOUZENUJEREIwODI5QzJFMiIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PriKAVcAAAMAUExURdbow/3+/d3swenz7dzsyMnj1eTx6sHgzvn79d3syuPw4trr29npzLXWxr/czt7txODt2dbozPr8+Nrq09npwcLexNnpxe318ebx0fT57eDtxfb68b3dytTo3LbZxdHlwbjYyPD26ubx1u724uHv5unz2NHlyfX59fz9+ujy3NnqyvD25cbh0sziwrjcxsrhwujy1sjgxdTmwfL49dzrvtrrvtHm28Dgzubw2uz03r3dxejy4r3bzMXgxdHlxOTwzdnq4s3k2Ony5t7t5sTf0tTmxbjZxvb6+Njpv9bov+HuzcnhyPL46dTmysTgycHgys/kwrzay+715dvqvc3jw97sytbo08Lez9rqx97u4ev03dvrxb7dw9nqv77czLrcyLrayvP47czjysHfx87kxdDm0uPvy+/28dDkztfox9box+rz2vH48trs4b7ey7rayNzrwcvkzPj7+brZyrjbxMrjzrbYyNzqycDezdfozs7l0c7kyvr8+8jizrrYyeDuy9Xnv7/czczjzvH35sbhzNbnzdPmz83kz7nayN/tzdXoyqvRv9/tyeDuyNvswcHgzbvcxNzrytrqy8HezsHfztXnzN7sydPmzNfoy9jpy9npy8Thzd7tyc7kzdTnzMDdzszkzdLmzNHlzdDlzeDuyc/kzcvjzcrjzdvqysnjzdvrysfizcbizcDczsPhzdLlzNTmzNLlzdvqy8jizcHdztboy9fpy8XhzcLhzcXizc3kzcDdz+Huyd7tyNzry93ry8Dcz9LlztPmzcHdz+DtytXnzbbXx/T569npysHez9/tyNTmzev04dzqy8fj0MHcz97sx8vjzNvrzOLuydTnzdzqzcDeztHmzsHcztboysjjzMniztPlzcrjzNrpycPgztbp4LnZydvsxtjpydXoyMXhzsrgw9fpyc/ky9PnzcjjzcnizMnjzNfptdzswNjpt9Hky+Lvysfiy8TgzMThy8TiztPlzNXmzev028Lgy8PgzMPhy9Dkx9XozeHuyP///xXNTX8AAB8KSURBVHja7N15dBzFnQdwzci6LGlkyYesy8i2jMeXfA2ObSQbX7KNhbGIA75vEDZYODbOcigOIAgk0bDWaR3WZUu+ZMknPoUJEASbACEh5NgcoJA4u5tdb9glm7DJTm33aGY0PdNdXVVd3TNd3fUHeS950eR98q3+/rq6R4oA5lK0IkwC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%3D%3D);");
                else
                    doc.Add("    background-image:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAugAAAGyCAMAAACBecGuAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAxBpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6QkMwREUwMUZEQUQ1MTFFMzg0NzE4OTJDOTc2NzhENTQiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6QkMwREUwMUVEQUQ1MTFFMzg0NzE4OTJDOTc2NzhENTQiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiBXaW5kb3dzIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9IkM4MTkwMzM1QzJGQUNGREQyNjE1Nzc5QjVGQkY3QjUyIiBzdFJlZjpkb2N1bWVudElEPSJDODE5MDMzNUMyRkFDRkREMjYxNTc3OUI1RkJGN0I1MiIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PsYwArgAAAG5UExURf3j4/3k5fz8/OXl5e7u7v3l5vLy8v709P3m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);");
            }
            doc.Add("    /*padding-top: 23px;*/");
            doc.Add("}");
            doc.Add("</style>");
            doc.Add(Resources.style);

            doc.Add("</head>");
            doc.Add("<body>");
            doc.Add("<!-- header here -->");
        }

        private void UpdateSelgerkoderUI()
        {
            selgerkodeList.Clear();
            comboDBselgerkode.Items.Clear();
            listBoxSk.Items.Clear();
            listBox_GraphSelgere.Items.Clear();
            if (!bwPopulateSk.IsBusy)
            {
                comboDBselgerkode.Items.Add("Laster..");
                listBoxSk.Items.Add("Laster..");
                listBox_GraphSelgere.Items.Add("Laster..");
                bwPopulateSk.RunWorkerAsync();
            }
        }

        private void bwPopulateSk_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            selgerkodeList.AddRange(sKoder.GetAlleSelgerkoder(appConfig.Avdeling));
        }

        private void bwPopulateSk_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            comboDBselgerkode.Items.Clear();
            comboDBselgerkode.Items.Add("ALLE");
            comboDBselgerkode.Items.AddRange(selgerkodeList.ToArray());
            if (comboDBselgerkode.Items.Count > 0)
                comboDBselgerkode.SelectedIndex = 0;

            listBoxSk.Items.Clear();
            listBoxSk.Items.AddRange(selgerkodeList.ToArray());

            listBox_GraphSelgere.Items.Clear();
            if (checkBox_GraphSort.Checked)
            {
                string[] str = selgerkodeList.ToArray();
                Array.Sort(str);
                listBox_GraphSelgere.Items.Clear();
                listBox_GraphSelgere.Items.Add("ALLE");
                listBox_GraphSelgere.Items.AddRange(str);
            }
            else
            {
                listBox_GraphSelgere.Items.Clear();
                listBox_GraphSelgere.Items.Add("ALLE");
                listBox_GraphSelgere.Items.AddRange(selgerkodeList.ToArray());
            }

            if (listBox_GraphSelgere.Items.Count > 0)
                listBox_GraphSelgere.SelectedIndex = 0;
        }

        private void NavigateWeb(WebBrowserNavigatingEventArgs e)
        {
            try
            {
                var url = e.Url.OriginalString;
                if (url.Contains("#link") && !url.Contains("#linkx"))
                {
                    var page = currentPage();
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
                else if (url.Contains("#vinn")) // klikk på vinnprodukt selgerkode
                {
                    int index = url.IndexOf("#");
                    var selger = url.Substring(index + 5, url.Length - index - 5);

                    if (selger.Length > 0)
                        RunVinnSelger(selger);
                    else
                        Logg.Alert("Ugyldig selgerkode!");
                }
                else if (url.Contains("#graf"))
                {
                    int index = url.IndexOf("#graf");
                    if (url.Length > index + 7)
                    {
                        string navn = url.Substring(index + 5, url.Length - index - 5);
                        RunGraph(navn);
                    }
                }
            }
            catch (Exception ex)
            {
                processing.SetVisible = false;
                FormError errorMsg = new FormError("Feil ved navigering", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void ReloadDatabase()
        {
            if (IsBusy(true))
                return;

            processing.SetVisible = true;
            processing.SetText = "Laster inn databasen på nytt..";
            processing.SetValue = 1;
            RetrieveDb(true);
            processing.SetValue = 30;
            RetrieveDbStore();
            RetrieveDbService();
            processing.SetValue = 40;
            Reload(true);
            processing.SetValue = 60;
            ReloadStore(true);
            processing.SetValue = 70;
            ReloadService(true);
            UpdateUi();
            processing.SetValue = 100;
            processing.SetText = "Ferdig";
            processing.HideDelayed();
            this.Activate();
        }

        delegate void SetStatusInfoCallback(string kat, string type, string str, DateTime date);

        private void SetStatusInfo(string kat, string type, string str, DateTime date)
        {
            if ((labelStatusTmerRanking.InvokeRequired || labelStatusTmerKveldstall.InvokeRequired
                || labelStatusTmerService.InvokeRequired || labelStatusTmerLager.InvokeRequired) && kat == "timer")
            {
                SetStatusInfoCallback d = new SetStatusInfoCallback(SetStatusInfo);
                this.Invoke(d, new object[] { kat, type, str });
            }
            else if ((labelStatusDbRanking.InvokeRequired || labelStatusDbLager.InvokeRequired || labelStatusDbService.InvokeRequired) && kat == "db")
            {
                SetStatusInfoCallback d = new SetStatusInfoCallback(SetStatusInfo);
                this.Invoke(d, new object[] { kat, type, str });
            }
            else
            {
                if (kat == "db")
                {
                    if (date == DateTime.MinValue)
                        str = "(" + type.Substring(0, 1).ToUpper() + type.Substring(1, type.Length - 1) + " databasen er tom)";

                    if (type == "ranking")
                        labelStatusDbRanking.Text = str;
                    else if (type == "lager")
                        labelStatusDbLager.Text = str;
                    else if (type == "service")
                        labelStatusDbService.Text = str;
                }
                if (kat == "timer")
                {
                    if (date == DateTime.MinValue)
                    {
                        str = "Avslått";
                        string stat = timerStatusLabelTimers.Text;
                        if (type == "ranking")
                            stat = stat.Replace("[Ranking]", "");
                        else if (type == "kveldstall")
                            stat = stat.Replace("[Kveld]", "");
                        else if (type == "lager")
                            stat = stat.Replace("[Lager]", "");
                        else if (type == "service")
                            stat = stat.Replace("[Service]", "");
                        timerStatusLabelTimers.Text = stat;
                    }
                    else
                    {
                        TimeSpan ts = date.Subtract(DateTime.Now);
                        str = str + date.ToShortDateString() + " " + date.ToShortTimeString() + " (om " + ToReadableString(ts) + ")";
                        string stat = timerStatusLabelTimers.Text;
                        if (type == "ranking" && !stat.Contains("[Ranking]"))
                            stat += "[Ranking]";
                        else if (type == "kveldstall" && !stat.Contains("[Kveld]"))
                            stat += "[Kveld]";
                        else if (type == "lager" && !stat.Contains("[Lager]"))
                            stat += "[Lager]";
                        else if (type == "service" && !stat.Contains("[Service]"))
                            stat += "[Service]";
                        timerStatusLabelTimers.Text = stat;
                    }

                    if (type == "ranking")
                        labelStatusTmerRanking.Text = str;
                    else if (type == "kveldstall")
                        labelStatusTmerKveldstall.Text = str;
                    else if (type == "lager")
                        labelStatusTmerLager.Text = str;
                    else if (type == "service")
                        labelStatusTmerService.Text = str;
                }
            }
        }

        public static string ToReadableString(TimeSpan span)
        {
            string formatted = string.Format("{0}{1}{2}",
                span.Duration().Days > 0 ? string.Format("{0:0} dag{1}, ", span.Days, span.Days == 1 ? String.Empty : "er") : string.Empty,
                span.Duration().Hours > 0 ? string.Format("{0:0} time{1}, ", span.Hours, span.Hours == 1 ? String.Empty : "r") : string.Empty,
                span.Duration().Minutes > 0 ? string.Format("{0:0} minutt{1}, ", span.Minutes, span.Minutes == 1 ? String.Empty : "er") : string.Empty);
                //span.Duration().Seconds > 0 ? string.Format("{0:0} second{1}", span.Seconds, span.Seconds == 1 ? String.Empty : "s") : string.Empty);

            if (formatted.EndsWith(", ")) formatted = formatted.Substring(0, formatted.Length - 2);

            if (string.IsNullOrEmpty(formatted)) formatted = "0 sekund";

            return formatted;
        }

        delegate void SetProgressStopCallback();

        public void ProgressStop()
        {
            if (statusStrip.InvokeRequired || buttonOppdater.InvokeRequired)
            {
                SetProgressStopCallback d = new SetProgressStopCallback(ProgressStop);
                this.Invoke(d, new object[] { });
            }
            else
            {
                if (!IsBusy(true, true))
                {
                    buttonOppdater.ForeColor = SystemColors.ControlText;
                    buttonOppdater.Text = "Oppdater";
                    progressbar.Style = ProgressBarStyle.Continuous;

                }
            }
        }

        delegate void SetProgressStartCallback();

        public void ProgressStart()
        {
            if (statusStrip.InvokeRequired || buttonOppdater.InvokeRequired)
            {
                SetProgressStartCallback d = new SetProgressStartCallback(ProgressStart);
                this.Invoke(d, new object[] { });
            }
            else
            {
                if (bwRanking.IsBusy || bwReport.IsBusy)
                {
                    buttonOppdater.Text = "Stop";
                    buttonOppdater.ForeColor = Color.Red;
                }
                progressbar.Style = ProgressBarStyle.Marquee;
            }
        }

        private void bwAutoRanking_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;
            var macroAttempt = 1;
            var macroMaxAttempts = 4;

            retrymacro:

            Logg.Log("Auto: Kjører makro.. [" + macroAttempt + "]");

            DateTime date = appConfig.dbTo;
            double span = (DateTime.Now - appConfig.dbTo).TotalDays;
            if (span > 31)
                date = DateTime.Now.AddMonths(-1); // Begrens oss til å importere en måned bak i tid
            if (appConfig.dbTo == DateTime.Now)
                date = GetFirstDayOfMonth(appConfig.dbTo);

            var macroForm = (FormMacro)StartMacro(date, macroProgram, bwAutoRanking, macroAttempt);
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

            // Oppdaterer gjeldene datoer.. vi oppdaterer UI etter ferdig jobb
            RetrieveDb(true);
            appConfig.strButikk = "";
            appConfig.strData = "";
            appConfig.strAudioVideo = "";
            appConfig.strTele = "";
            appConfig.strOversikt = "";
            appConfig.strObsolete = "";

            // Hvis valgt, avbryt utsending hvis vi ikke har tall fra i går.
            if (appConfig.epostOnlySendUpdated)
            {
                int dager = (DateTime.Now - appConfig.dbTo).Days;
                Logg.Log("Auto: Kontrollerer dato på sist transaksjon.. (dager: " + dager + ")");
                if (dager > 1)
                {
                    Logg.Log("Auto: Fant ingen transaksjoner fra gårsdagen. Avbryter..", Color.Red);
                    return;
                }
                else
                    Logg.Log("Auto: Fant transaksjoner fra gårsdagen. Fortsetter..", Color.Green);
            }

            // Legg ved service rapport
            if (appConfig.AutoService && appConfig.epostIncService && service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
            {
                Logg.Log("Auto: Legger ved service rapport..");
                MakeServiceOversikt(true, bwAutoRanking);
            }

            if (appConfig.onlineImporterAuto)
            {
                OnlineImporter importer = new OnlineImporter(this);

                Logg.Log("Auto: Oppdatere Prisguide.no produkter..");
                processing.SetText = "Oppdaterer Prisguide.no produkter..";

                bool successfull = importer.StartProcessingPrisguide(bwAutoStore);
                if (!successfull)
                    Logg.Log("Auto: Misslykket oppdatering av Prisguide.no produkter. Se logg for detaljer.", Color.Red);

                Logg.Log("Auto: Oppdatere Ukenytt produkter fra Elkjop.no..");
                processing.SetText = "Oppdaterer Ukenytt produkter fra Elkjop.no..";

                successfull = importer.StartProcessingWeekly(bwAutoStore);
                if (!successfull)
                    Logg.Log("Auto: Misslykket oppdatering av Ukenytt produkter. Se logg for detaljer.", Color.Red);
            }

            processing.SetText = "Klargjør for sending av ranking..";
            Logg.Log("Auto: Konverterer ranking til PDF..");

            KgsaEmail kgsaEmail = new KgsaEmail(this);

            string pdfAlt = CreatePDF("Full", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'Full'..";
            kgsaEmail.Send(pdfAlt, appConfig.dbTo, "Full", appConfig.epostEmne, appConfig.epostBody);

            string pdfComputer = CreatePDF("Computer", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'Computer'..";
            kgsaEmail.Send(pdfComputer, appConfig.dbTo, "Computer", appConfig.epostEmne, appConfig.epostBody);

            string pdfAudioVideo = CreatePDF("AudioVideo", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'AudioVideo'..";
            kgsaEmail.Send(pdfAudioVideo, appConfig.dbTo, "AudioVideo", appConfig.epostEmne, appConfig.epostBody);

            string pdfTelecom = CreatePDF("Telecom", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'Telecom'..";
            kgsaEmail.Send(pdfTelecom, appConfig.dbTo, "Telecom", appConfig.epostEmne, appConfig.epostBody);

            string pdfCross = CreatePDF("Cross", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'Cross'..";
            kgsaEmail.Send(pdfCross, appConfig.dbTo, "Cross", appConfig.epostEmne, appConfig.epostBody);

            string pdfVinn = CreatePDF("Vinnprodukter", "", bwAutoRanking);
            processing.SetText = "Sender e-post for gruppe 'Vinn'..";
            kgsaEmail.Send(pdfVinn, appConfig.dbTo, "Vinnprodukter", appConfig.epostEmne, appConfig.epostBody);

            if (appConfig.ignoreSunday && DateTime.Now.Date.DayOfWeek == DayOfWeek.Sunday)
            {
                string pdfAltUkestart = CreatePDF("FullUkestart", "", bwAutoRanking);
                processing.SetText = "Sender e-post for gruppe 'FullUkestart'..";
                kgsaEmail.Send(pdfAltUkestart, appConfig.dbTo, "FullUkestart", appConfig.epostEmne, appConfig.epostBody);
            }
            else if (!appConfig.ignoreSunday && DateTime.Now.Date.DayOfWeek == DayOfWeek.Monday)
            {
                string pdfAltUkestart = CreatePDF("FullUkestart", "", bwAutoRanking);
                processing.SetText = "Sender e-post for gruppe 'FullUkestart'..";
                kgsaEmail.Send(pdfAltUkestart, appConfig.dbTo, "FullUkestart", appConfig.epostEmne, appConfig.epostBody);
            }

            if (String.IsNullOrEmpty(pdfAlt) && String.IsNullOrEmpty(pdfComputer) && String.IsNullOrEmpty(pdfAudioVideo) && String.IsNullOrEmpty(pdfTelecom))
            {
                // Feil under pdf generering og sending av epost
                Logg.Log("Auto: Feil oppstod under generering av PDF og sending av epost. Se logg for detaljer.", Color.Red);
                return;
            }

            if (appConfig.pdfExport && !String.IsNullOrEmpty(appConfig.pdfExportFolder) && !String.IsNullOrEmpty(pdfAlt))
            {
                Logg.Log("Auto: Lagrer kopi av PDF til..");
                try
                {
                    File.Copy(pdfAlt, appConfig.pdfExportFolder);
                }
                catch { Logg.Log("Auto: Feil oppstod under kopiering av PDF."); }
            }

            // App database update
            if (appConfig.blueProductAutoUpdate)
            {
                Logg.Log("Auto: Oppdaterer produktdata databasen for App..");
                AppManager appMng = new AppManager(this);
                appMng.ExportProductDatabase(bwAutoRanking);
            }
        }

        private void bwAutoRanking_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;
            RetrieveDb();
            RetrieveDbService();
            RetrieveDbStore();
            Reload(true);
            ReloadService();
            UpdateUi();
            Logg.Log("Auto: Avsluttet.");
            processing.HideDelayed();
            this.Activate();
        }

        private void delayedAutoRanking()
        {
            try
            {
                processing.SetVisible = true;
                processing.SetProgressStyle = ProgressBarStyle.Continuous;
                processing.SetBackgroundWorker = bwAutoRanking;
                for (int b = 0; b < 100; b++)
                {
                    processing.SetText = "Starter automatisk ranking om " + (((b / 10) * -1) + 10) + " sekunder..";
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

                processing.SetProgressStyle = ProgressBarStyle.Marquee;
                processing.SetText = "Starter automatisk ranking rutine..";
                Logg.Log("Starter automatisk ranking rutine..");
                processing.SetBackgroundWorker = bwAutoRanking;
                bwAutoRanking.RunWorkerAsync(); // Starter jobb som starter makro, importering, ranking, pdf konvertering og sending på mail.
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Uhåndtert unntak oppstod ved delayedAutoRanking.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void delayedAutoStore()
        {
            try
            {
                processing.SetVisible = true;
                processing.SetProgressStyle = ProgressBarStyle.Continuous;
                processing.SetBackgroundWorker = bwAutoStore;
                for (int b = 0; b < 100; b++)
                {
                    processing.SetText = "Starter automatisk lager import om " + (((b / 10) * -1) + 10) + " sekunder..";
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

                RunAutoObsoleteImport();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwAutoStore_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;
            var macroAttempt = 1;
            var macroMaxAttempts = 4;

            retrymacro:

            Logg.Log("Auto: Kjører makro.. [" + macroAttempt + "]");

            FormMacro formMacro = (FormMacro)StartMacro(DateTime.Now.AddDays(-1), macroProgramStore, bwAutoStore, macroAttempt);
            e.Result = formMacro.errorCode;
            if (formMacro.errorCode != 0)
            {
                // Feil oppstod under kjøring av macro
                macroAttempt++;
                if (formMacro.errorCode != 2)
                {
                    Logg.Log("Auto: Feil oppstod under kjøring av makro. Feilbeskjed: " + formMacro.errorMessage + " Kode: " + formMacro.errorCode, Color.Red);
                    processing.SetText = formMacro.errorMessage;
                }

                for (int i = 0; i < 60; i++)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
                if (formMacro.errorCode == 6)
                    return;
                if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                    goto retrymacro;

                e.Result = formMacro.errorCode;
                return;
            }

            if (!File.Exists(appConfig.csvElguideExportFolder + @"\wobsolete.zip") || File.GetLastWriteTime(appConfig.csvElguideExportFolder + @"\wobsolete.zip").Date != DateTime.Now.Date)
            {
                // CSV finnes ikke eller filen er ikke oppdatert i dag i.e. data ble ikke eksportert riktig med makro
                macroAttempt++;
                Logg.Log("Auto: CSV er IKKE oppdatert, eller ingen tilgang. Sjekk CSV lokasjon og makro innstillinger.", Color.Red);
                for (int i = 0; i < 60; i++)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
                if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                    goto retrymacro;

                return;
            }

            Logg.Log("Auto: Importerer data..");
            string extracted = obsolete.Decompress(appConfig.csvElguideExportFolder + @"\wobsolete.zip");
            if (!String.IsNullOrEmpty(extracted))
            {
                if (obsolete.Import(extracted, bwImportObsolete))
                    e.Result = 0;
                else
                {
                    Logg.Log("Auto: Importering mislyktes!", Color.Red);
                    macroAttempt++;
                    if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                        goto retrymacro;

                    e.Result = 6;
                    return;
                }
            }
            else
            {
                Logg.Log("Auto: Utpakking mislykket av arkiv.", Color.Red);
                if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                    goto retrymacro;

                e.Result = 6;
                return;
            }

            if (appConfig.blueInventoryAutoUpdate)
            {
                Logg.Log("Auto: Oppdaterer varebeholdnings databasen for App..");
                processing.SetText = "Oppdaterer varebeholdning for App..";
                AppManager appMng = new AppManager(this);
                if (appMng.ImportAndConvertInventory(bwAutoStore))
                {
                    if (File.Exists(BluetoothServer.inventoryFilename))
                    {
                        appConfig.blueInventoryLastDate = DateTime.Now;
                        appConfig.blueInventoryReady = true;
                        Logg.Log("Auto: Varebeholdnings databasen er klar for App.");
                    }
                    else
                        appConfig.blueInventoryReady = false;
                }
                else
                    appConfig.blueInventoryReady = false;
            }

            if (appConfig.onlineImporterAuto)
            {
                OnlineImporter importer = new OnlineImporter(this);

                Logg.Log("Auto: Oppdatere Prisguide.no produkter..");
                processing.SetText = "Oppdaterer Prisguide.no produkter..";

                bool successfull = importer.StartProcessingPrisguide(bwAutoStore);
                if (!successfull)
                    Logg.Log("Auto: Misslykket oppdatering av Prisguide.no produkter. Se logg for detaljer.", Color.Red);

                Logg.Log("Auto: Oppdatere Ukenytt produkter fra Elkjop.no..");
                processing.SetText = "Oppdaterer Ukenytt produkter fra Elkjop.no..";

                successfull = importer.StartProcessingWeekly(bwAutoStore);
                if (!successfull)
                    Logg.Log("Auto: Misslykket oppdatering av Ukenytt produkter. Se logg for detaljer.", Color.Red);
            }

        }

        private void bwAutoStore_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;
            if (e.Cancelled)
            {
                Logg.Log("Auto: Jobb avbrutt av bruker.");
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
                    database.ClearCacheTables();
                    RetrieveDb();
                    RetrieveDbService();
                    RetrieveDbStore();
                    Reload(true);
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

        public void RunMakePDF(bool all = false)
        {
            if (!bwPDF.IsBusy)
            {
                processing.SetVisible = true;
                processing.SetText = "Konverterer til PDF..";
                bwPDF.RunWorkerAsync(all);
                while (bwPDF.IsBusy)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
                processing.HideDelayed();
                this.Activate();
            }
        }

        public void RunMakeBudgetPDF(bool all = false)
        {
            if (!bwBudgetPDF.IsBusy)
            {
                processing.SetVisible = true;
                processing.SetText = "Konverterer budsjett til PDF..";
                bwBudgetPDF.RunWorkerAsync(all);
                while (bwBudgetPDF.IsBusy)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
                processing.HideDelayed();
                this.Activate();
            }
        }

        public void RunOpenPDF()
        {
            if (!bwOpenPDF.IsBusy)
            {
                processing.SetVisible = true;
                processing.SetText = "Konverterer til PDF..";
                bwOpenPDF.RunWorkerAsync();
            }
        }

        private void bwOpenPDF_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            e.Result = CreatePDF("", currentPageFile(), bwPDF, true);
        }

        private void bwOpenPDF_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            string result = (string)e.Result;
            if (result.Length > 3)
            {
                processing.SetText = "Åpner PDF..";
                Logg.Log("Åpner.. file://" + result.Replace(' ', (char)160));
                try
                {
                    System.Diagnostics.Process.Start(result);
                }
                catch(FileNotFoundException ex)
                {
                    Logg.Unhandled(ex);
                    Logg.Log("Kunne ikke åpne PDF. Se logg for detaljer.", Color.Red);
                }
                ProgressStop();
                processing.SetVisible = false;
            }
            else
            {
                ProgressStop();
                processing.HideDelayed();
                this.Activate();
            }
        }

        private void bwPDF_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            bool value = (bool)e.Argument;

            if (value)
                filenamePDF = CreatePDF("Full", "", bwPDF);
            else
                filenamePDF = CreatePDF("", currentPageFile(), bwPDF);
        }

        private void bwBudgetPDF_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            bool value = (bool)e.Argument;

            if (value)
                filenamePDF = CreatePDF("FullBudget", "", bwBudgetPDF);
            else
                filenamePDF = CreatePDF("", currentBudgetPageFile(), bwBudgetPDF);
        }

        private void bwPDF_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true, true))
            {
                processing.SetText = "Ferdig!";
                processing.HideDelayed();
            }
        }

        private void savePDF(bool all = false)
        {
            try
            {
                if (!EmptyDatabase() || !EmptyStoreDatabase())
                {
                    SaveFileDialog SD = new SaveFileDialog();
                    SD.Filter = "PDF Filer (*.pdf)|*.pdf|Vis alle filer (*.*)|*.*";

                    string file = "";
                    if (all)
                        file = "KGSA " + appConfig.Avdeling + " " + pickerRankingDate.Value.ToString("yyy-MM-dd") + ".pdf";
                    else
                        file = "KGSA " + appConfig.Avdeling + " " + currentPage() + " " + pickerRankingDate.Value.ToString("yyy-MM-dd") + ".pdf";
                    SD.FileName = file;
                    if (all)
                        SD.Title = "Lagre Ranking PDF som";
                    else
                        SD.Title = "Lagre [" + currentPage() + "] PDF som";
                    if (SD.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            RunMakePDF(all);
                            if (!String.IsNullOrEmpty(filenamePDF))
                            {
                                Logg.Log("PDF ferdig generert.");
                                if (File.Exists(SD.FileName))
                                {
                                    File.Delete(SD.FileName);
                                    Logg.Log("Skriver over eksisterende fil..");
                                }
                                File.Copy(filenamePDF, SD.FileName);
                                Logg.Log("Fullført lagring av PDF. file://" + SD.FileName.Replace(' ', (char)160), Color.Green);
                            }
                            else
                            {
                                Logg.Log("Avbrutt.", Color.Red);
                                return;
                            }
                        }
                        catch (IOException)
                        {
                            MessageBox.Show("KGSA - Feil", "Filen er i bruk eller ble nektet tilgang.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Logg.Log("Filen er i bruk eller ingen tilgang.", Color.Red);
                        }
                        catch
                        {
                            Logg.Log("Ukjent feil oppstod under eksportering av databasen. Operasjon avbrutt.", Color.Red);
                        }
                    }
                    else
                        Logg.Log("Avbrutt.");
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void saveBudgetPDF(bool all = false)
        {
            try
            {
                if (!EmptyDatabase())
                {
                    SaveFileDialog SD = new SaveFileDialog();
                    SD.Filter = "PDF Filer (*.pdf)|*.pdf|Vis alle filer (*.*)|*.*";

                    string file = "";
                    if (all)
                        file = "KGSA Budsjett " + appConfig.Avdeling + " " + pickerRankingDate.Value.ToString("yyy-MM-dd") + ".pdf";
                    else
                        file = "KGSA Budsjett " + appConfig.Avdeling + " " + currentBudgetPage() + " " + pickerRankingDate.Value.ToString("yyy-MM-dd") + ".pdf";
                    SD.FileName = file;
                    if (all)
                        SD.Title = "Lagre Budsjett PDF som";
                    else
                        SD.Title = "Lagre [" + currentPage() + "] PDF som";
                    if (SD.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            RunMakeBudgetPDF(all);
                            if (!String.IsNullOrEmpty(filenamePDF))
                            {
                                Logg.Log("PDF ferdig generert.");
                                if (File.Exists(SD.FileName))
                                {
                                    File.Delete(SD.FileName);
                                    Logg.Log("Skriver over eksisterende fil..");
                                }
                                File.Copy(filenamePDF, SD.FileName);
                                Logg.Log("Fullført lagring av PDF. (" + SD.FileName + ")", Color.Green);
                            }
                            else
                            {
                                Logg.Log("Avbrutt.", Color.Red);
                                return;
                            }
                        }
                        catch (IOException)
                        {
                            MessageBox.Show("KGSA - Feil", "Filen er i bruk eller ble nektet tilgang.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Logg.Log("Filen er i bruk eller ingen tilgang.", Color.Red);
                        }
                        catch
                        {
                            Logg.Log("Ukjent feil oppstod under eksportering av databasen. Operasjon avbrutt.", Color.Red);
                        }
                    }
                    else
                        Logg.Log("Avbrutt.");
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void TransExportCSV()
        {
            try
            {
                if (dataGridTransaksjoner.DataSource == null)
                {
                    Logg.Log("Transaksjons tabellen er tom!", Color.Red);
                    return;
                }
                Logg.Status("Forbereder..");
                DataTable dt = GetContentAsDataTable();

                if (dt.Rows.Count > 0)
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                    sfd.FilterIndex = 2;
                    sfd.RestoreDirectory = true;
                    sfd.FileName = "KGSA " + appConfig.Avdeling + " " + pickerDBTil.Value.ToString("yyy-MM-dd") + " Transaksjoner.csv";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Logg.Log("Lagrer CSV..");
                        string filnavn = MakeCSV(dt, "Trans", "Custom", pickerDBTil.Value);
                        File.Copy(filnavn, sfd.FileName, true);
                        Logg.Log("Fullført eksportering av CSV! (" + sfd.FileName + ")", Color.Green);
                    }
               }
               else
                    Logg.Log("Transaksjons tabellen er tom!", Color.Red);
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private string MakeCSV(DataTable dt, string kat, string caption, DateTime dato)
        {
            try
            {
                string filnavn = "";

                StringBuilder sb = new StringBuilder();

                IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName);
                sb.AppendLine(string.Join(";", columnNames.ToArray()));

                foreach (DataRow row in dt.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                    sb.AppendLine(string.Join(";", fields.ToArray()));
                }

                filnavn = System.IO.Path.GetTempPath() + "KGSA " + appConfig.Avdeling + " " + dato.ToString("yyy-MM-dd") + " " + kat + " " + caption + ".csv";

                File.WriteAllText(filnavn, sb.ToString());

                return filnavn;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private DataTable GetContentAsDataTable(bool IgnoreHideColumns = false)
        {
            try
            {
                if (dataGridTransaksjoner.ColumnCount == 0) return null;
                DataTable dtSource = new DataTable();
                foreach (DataGridViewColumn col in dataGridTransaksjoner.Columns)
                {
                    if (IgnoreHideColumns & !col.Visible) continue;
                    if (String.IsNullOrEmpty(col.Name)) continue;
                    dtSource.Columns.Add(col.Name, col.ValueType);
                    dtSource.Columns[col.Name].Caption = col.HeaderText;
                }
                if (dtSource.Columns.Count == 0) return null;
                foreach (DataGridViewRow row in dataGridTransaksjoner.Rows)
                {
                    DataRow drNewRow = dtSource.NewRow();
                    foreach (DataColumn col in dtSource.Columns)
                    {
                        drNewRow[col.ColumnName] = row.Cells[col.ColumnName].Value;
                    }
                    dtSource.Rows.Add(drNewRow);
                }
                return dtSource;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        delegate void ChangeRankDateTimePickerCallback(DateTime date, DateTime dateMin, DateTime dateMax);
        private void ChangeRankDateTimePicker(DateTime date, DateTime dateMin, DateTime dateMax)
        {
            try
            {
                if (pickerRankingDate.InvokeRequired)
                {
                    ChangeRankDateTimePickerCallback d = new ChangeRankDateTimePickerCallback(ChangeRankDateTimePicker);
                    this.Invoke(d, new object[] { date, dateMin, dateMax });
                }
                else
                {
                    if (dateMin > dateMax)
                        dateMin = dateMax;

                    if (pickerRankingDate.MaxDate > dateMin)
                    {
                        pickerRankingDate.MinDate = rangeMin;
                        pickerRankingDate.MaxDate = rangeMax;
                    }

                    pickerRankingDate.MaxDate = dateMax;
                    pickerRankingDate.MinDate = dateMin;

                    if (date.Date >= dateMin.Date && date.Date <= dateMax.Date)
                        pickerRankingDate.Value = date;
                    else
                        pickerRankingDate.Value = dateMax;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        delegate void ChangeBudgetDateTimePickerCallback(DateTime date, DateTime dateMin, DateTime dateMax);
        private void ChangeBudgetDateTimePicker(DateTime date, DateTime dateMin, DateTime dateMax)
        {
            try
            {
                if (pickerBudget.InvokeRequired)
                {
                    ChangeBudgetDateTimePickerCallback d = new ChangeBudgetDateTimePickerCallback(ChangeBudgetDateTimePicker);
                    this.Invoke(d, new object[] { date, dateMin, dateMax });
                }
                else
                {
                    if (dateMin > dateMax)
                        dateMin = dateMax;

                    if (pickerBudget.MaxDate > dateMin)
                    {
                        pickerBudget.MinDate = rangeMin;
                        pickerBudget.MaxDate = rangeMax;
                    }

                    pickerBudget.MaxDate = dateMax;
                    pickerBudget.MinDate = dateMin;

                    if (date.Date >= dateMin.Date && date.Date <= dateMax.Date)
                        pickerBudget.Value = date;
                    else
                        pickerBudget.Value = dateMax;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        delegate void ChangeServiceDateTimePickerCallback(DateTime date, DateTime dateMin, DateTime dateMax);
        private void ChangeServiceDateTimePicker(DateTime date, DateTime dateMin, DateTime dateMax)
        {
            if (pickerServiceDato.InvokeRequired)
            {
                ChangeServiceDateTimePickerCallback d = new ChangeServiceDateTimePickerCallback(ChangeServiceDateTimePicker);
                this.Invoke(d, new object[] { date, dateMin, dateMax });
            }
            else
            {
                if (pickerServiceDato.MaxDate > dateMin)
                {
                    pickerServiceDato.MinDate = rangeMin;
                    pickerServiceDato.MaxDate = rangeMax;
                }

                pickerServiceDato.MaxDate = dateMax;
                pickerServiceDato.MinDate = dateMin;

                if (date.Date >= dateMin.Date && date.Date <= dateMax.Date)
                    pickerServiceDato.Value = date;
                else
                    pickerServiceDato.Value = dateMax;
            }
        }

        delegate void ChangeStoreDateTimePickerCallback(DateTime date, DateTime dateMin, DateTime dateMax);
        private void ChangeStoreDateTimePicker(DateTime date, DateTime dateMin, DateTime dateMax)
        {
            if (pickerLagerDato.InvokeRequired)
            {
                ChangeStoreDateTimePickerCallback d = new ChangeStoreDateTimePickerCallback(ChangeStoreDateTimePicker);
                this.Invoke(d, new object[] { date, dateMin, dateMax });
            }
            else
            {
                if (pickerLagerDato.MaxDate > dateMin)
                {
                    pickerLagerDato.MinDate = rangeMin;
                    pickerLagerDato.MaxDate = rangeMax;
                }

                pickerLagerDato.MaxDate = dateMax;
                pickerLagerDato.MinDate = dateMin;

                if (date.Date >= dateMin.Date && date.Date <= dateMax.Date)
                    pickerLagerDato.Value = date;
                else
                    pickerLagerDato.Value = dateMax;
            }
        }

        private void ProgressReport(decimal current, StatusProgress status)
        {
            try
            {
                decimal total = status.total;
                decimal percent = 0;

                if (total != 0)
                    percent = current / total;
                decimal gap = status.end - status.start;

                if (gap != 0)
                {
                    decimal gapValue = gap * percent;

                    if ((int)(status.start + gapValue) >= 0 && (int)(status.start + gapValue) <= 100)
                        processing.SetValue = (int)(status.start + gapValue);
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }


        delegate void SetLogCallback(KgsaLog logArg);

        void LogMessage_LogAdded(object sender, EventArgs e)
        {
            try
            {
                if (this.richLog.InvokeRequired || this.statusStrip.InvokeRequired)
                {
                    if (!IsHandleCreated)
                        return;

                    SetLogCallback d = new SetLogCallback(InvokeLogger);
                    this.Invoke(d, new object[] { Logg.GetLastLog() });
                }
                else
                {
                    InvokeLogger(Logg.GetLastLog());
                }
            }
            catch (Exception ex) { Console.WriteLine("Exception at LogMessage_LogAdded() " + ex.Message); }
        }

        private void AddToLoggCache(string line)
        {
            loggCacheMessages.Add(line);
            if (loggCacheMessages.Count > 10)
                loggCacheMessages.RemoveAt(0);
        }

        public static List<string> loggCacheMessages = new List<string>() { };

        private void InvokeLogger(KgsaLog logArg)
        {
            try
            {
                var log = (KgsaLog)logArg;
                if (log.debug && !appConfig.debug) // hvis det er en debug melding og debug er ikke aktivert; ikke send melding!
                    return;

                string str = log.message;
                str = str.Trim();
                str = str.Replace("\n", " ");
                str = str.Replace("\r", String.Empty);
                str = str.Replace("\t", String.Empty);

                if (!log.fileonly && !log.logonly && !log.debug)
                    ClearMessageTimer();

                var file = new StreamWriter(settingsPath + @"\Log.txt", true, Encoding.Unicode);
                DateTime t = DateTime.Now;
                string logLine = "";
                if (log.debug)
                    logLine = t.ToShortDateString() + " - " + t.ToShortTimeString() + ":" + t.ToString("ss") + " (debug) : " + str;
                else
                    logLine = t.ToShortDateString() + " - " + t.ToShortTimeString() + ":" + t.ToString("ss") + " : " + str;
                if (!log.statusonly)
                    file.WriteLine(logLine);
                file.Close();
                if (!log.debug)
                    AddToLoggCache(t.ToShortTimeString() + ":" + t.ToString("ss") + " : " + str); // Lagre til cache, for bruk til async stuff

                if (!log.fileonly)
                {
                    if (!log.logonly)
                    {
                        toolStripStatusLabel1.ForeColor = log.color;
                        toolStripStatusLabel1.Text = str;
                        statusStrip.Refresh();
                    }

                    if (!log.statusonly)
                    {
                        richLog.AppendText(logLine + Environment.NewLine, log.color);
                        richLog.ScrollToCaret();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kritisk feil i logg system.\n" + ex.Message, "KGSA - Feil", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static DateTime datoPeriodeFra = DateTime.MinValue;
        public static DateTime datoPeriodeTil;
        public static bool datoPeriodeVelger;
        private void velgDato()
        {
            try
            {
                if (EmptyDatabase())
                    return;

                if (datoPeriodeFra == default(DateTime))
                    datoPeriodeFra = appConfig.dbTo; // Endret fra dbFraDT
                if (datoPeriodeTil == default(DateTime))
                    datoPeriodeTil = appConfig.dbTo;

                var datovelger = new VelgPeriode(this);
                if (datovelger.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    datoPeriodeFra = datovelger.dtFra;
                    datoPeriodeTil = datovelger.dtTil;
                    if (datoPeriodeFra.Date != datoPeriodeTil.Date)
                    {
                        datoPeriodeVelger = true;

                        panelNotification.Visible = true;

                        panelGraphNotification.Visible = true;

                        labelNotificationText.Text = "Periode: " + datoPeriodeFra.ToString("dddd d. MMMM yyyy", norway) + " - " + datoPeriodeTil.ToString("dddd d. MMMM yyyy", norway);

                        labelGraphNotificationText.Text = "Periode: " + datoPeriodeFra.ToString("dddd d. MMMM yyyy", norway) + " - " + datoPeriodeTil.ToString("dddd d. MMMM yyyy", norway);

                        if (tabControlMain.SelectedTab == tabPageRank)
                            UpdateRank();
                        else if (tabControlMain.SelectedTab == tabPageGrafikk)
                        {
                            if (_graphInitialized)
                                UpdateGraph();
                            else
                                InitGraph();
                        }
                    }
                    else
                        Logg.Log("Fra og til dato var helt lik!", Color.Red);
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void UpdateRank(string katArg = "")
        {
            string page = currentPage();
            if (String.IsNullOrEmpty(katArg) && !String.IsNullOrEmpty(page))
                ClearHash(page);
            if (!String.IsNullOrEmpty(katArg)) // Har fått instruks om å oppdatere en spesifikk side..
                RunRanking(katArg);
            else
            {
                // Har ikke fått instruks om hvilken side, sjekker hvilken side som er synlig..
                if (!String.IsNullOrEmpty(page)) // Har funnet hvilken side vi er på
                    RunRanking(page);
                else if (!String.IsNullOrEmpty(savedPage)) // ..eller sjekk siste åpnet.
                    RunRanking(savedPage);
                else if (!String.IsNullOrEmpty(appConfig.savedPage)) // eller kanskje den er lagret?
                    RunRanking(appConfig.savedPage);
                else // OK, vi gir opp og åpner standard side..
                    RunRanking("Data");
            }
        }

        private void UpdateBudget(BudgetCategory cat = BudgetCategory.None)
        {
            if (cat != BudgetCategory.None)
                ClearBudgetHash(cat);

            if (cat != BudgetCategory.None)
                RunBudget(cat);
            else
            {
                BudgetCategory page = currentBudgetPage();

                if (page != BudgetCategory.None)
                    ClearBudgetHash(page);

                if (page != BudgetCategory.None)
                    RunBudget(page);
                else if (savedBudgetPage != BudgetCategory.None)
                    RunBudget(savedBudgetPage);
                else if (appConfig.savedBudgetPage != BudgetCategory.None)
                    RunBudget(appConfig.savedBudgetPage);
                else
                    RunBudget(BudgetCategory.MDA);
            }
        }

        public bool readControlState()
        {
            if (tabControlMain.InvokeRequired)
            {
                return (bool)tabControlMain.Invoke(
                  new Func<Boolean>(() => readControlState())
                );
            }
            else
            {
                bool curTab = tabControlMain.SelectedTab == tabPageService;
                return curTab;
            }
        }

        public string readCurrentTab()
        {
            if (tabControlMain.InvokeRequired)
            {
                return (string)tabControlMain.Invoke(
                  new Func<string>(() => readCurrentTab())
                );
            }
            else
            {
                if (tabControlMain.SelectedTab == tabPageService)
                    return "Service";
                else if (tabControlMain.SelectedTab == tabPageStore)
                    return "Store";
                else if (tabControlMain.SelectedTab == tabPageBudget)
                    return "Budget";
                else if (tabControlMain.SelectedTab == tabPageRank)
                    return "Ranking";
                else if (tabControlMain.SelectedTab == tabPageLog)
                    return "Log";
                else
                    return "";
            }
        }


        public string currentPage()
        {
            try
            {
                string curTab = readCurrentTab();
                if (curTab == "Service" && webService.Url != null)
                {
                    string str = webService.Url.OriginalString;
                    if (str.Contains("serviceOversikt.html"))
                        return "ServiceOversikt";
                    else if (str.Contains("serviceList.html"))
                        return "ServiceList";
                    else if (str.Contains("serviceDetails.html"))
                        return "ServiceDetails";
                    else
                        return "";
                }
                else if (curTab == "Budget" && webBudget.Url != null)
                {
                    string str = webBudget.Url.OriginalString;
                    if (str.Contains("budsjettMda.html"))
                        return "MDA";
                    else if (str.Contains("budsjettAudioVideo.html"))
                        return "AudioVideo";
                    else if (str.Contains("budsjettSda.html"))
                        return "SDA";
                    else if (str.Contains("budsjettTele.html"))
                        return "Tele";
                    else if (str.Contains("budsjettData.html"))
                        return "Data";
                    else if (str.Contains("budsjettCross.html"))
                        return "Cross";
                    else if (str.Contains("budsjettKasse.html"))
                        return "Kasse";
                    else if (str.Contains("budsjettAftersales.html"))
                        return "Aftersales";
                    else if (str.Contains("budsjettMdaSda.html"))
                        return "MDASDA";
                    else if (str.Contains("budsjettButikk.html"))
                        return "Butikk";
                    else if (str.Contains("budsjettDaglig.html"))
                        return "Daglig";
                    else
                        return "";
                }
                else if (curTab == "Ranking" && webHTML.Url != null)
                {
                    string str = webHTML.Url.OriginalString;
                    if (str.Contains("rankingButikk.html"))
                        return "Butikk";
                    else if (str.Contains("rankingKnowHow.html"))
                        return "KnowHow";
                    else if (str.Contains("rankingData.html"))
                        return "Data";
                    else if (str.Contains("rankingAudioVideo.html"))
                        return "AudioVideo";
                    else if (str.Contains("rankingTele.html"))
                        return "Tele";
                    else if (str.Contains("rankingOversikt.html"))
                        return "Oversikt";
                    else if (str.Contains("rankingBudsjett.html"))
                        return "Budsjett";
                    else if (str.Contains("rankingToppselgere.html"))
                        return "Toppselgere";
                    else if (str.Contains("rankingLister.html"))
                        return "Lister";
                    else if (str.Contains("rankingRapport.html"))
                        return "Rapport";
                    else if (str.Contains("rankingQuick.html"))
                        return "Quick";
                    else if (str.Contains("rankingPeriode.html"))
                        return "Periode";
                    else if (str.Contains("rankingGraf.html"))
                        return "Graf";
                    else if (str.Contains("rankingVinnprodukter.html"))
                        return "Vinnprodukter";
                    else if (str.Contains("rankingVinnprodukterSelger.html"))
                        return "VinnSelger";
                    else if (str.Contains("rankingAvdTjenester.html"))
                        return "Tjenester";
                    else if (str.Contains("rankingAvdSnittpriser.html"))
                        return "Snittpriser";
                    else
                        return "";
                }
                else if (curTab == "Store" && webLager.Url != null)
                {
                    string str = webLager.Url.OriginalString;
                    if (str.Contains("storeLagerstatus.html"))
                        return "Obsolete";
                    else if (str.Contains("storeObsoleteList.html"))
                        return "ObsoleteList";
                    else if (str.Contains("storeMarginer.html"))
                        return "LagerMarginer";
                    else if (str.Contains("storeWeekly.html"))
                        return "LagerUkeAnnonser";
                    else if (str.Contains("storeWeeklyOverview.html"))
                        return "LagerUkeAnnonserOversikt";
                    else if (str.Contains("storePrisguide.html"))
                        return "LagerPrisguide";
                    else if (str.Contains("storePrisguideOverview.html"))
                        return "LagerPrisguideOversikt";
                    else
                        return "";
                }
                else
                    return "";
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private BudgetCategory currentBudgetPage()
        {
            try
            {
                if (webBudget.Url != null)
                {
                    string str = webBudget.Url.OriginalString;
                    if (str.Contains("budsjettMda.html"))
                        return BudgetCategory.MDA;
                    else if (str.Contains("budsjettAudioVideo.html"))
                        return BudgetCategory.AudioVideo;
                    else if (str.Contains("budsjettSda.html"))
                        return BudgetCategory.SDA;
                    else if (str.Contains("budsjettTele.html"))
                        return BudgetCategory.Tele;
                    else if (str.Contains("budsjettData.html"))
                        return BudgetCategory.Data;
                    else if (str.Contains("budsjettCross.html"))
                        return BudgetCategory.Cross;
                    else if (str.Contains("budsjettKasse.html"))
                        return BudgetCategory.Kasse;
                    else if (str.Contains("budsjettAftersales.html"))
                        return BudgetCategory.Aftersales;
                    else if (str.Contains("budsjettDaily.html"))
                        return BudgetCategory.Aftersales;
                    else if (str.Contains("budsjettDaglig.html"))
                        return BudgetCategory.Daglig;
                }
                return BudgetCategory.None;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return BudgetCategory.None;
            }
        }

        private string currentPageFile()
        {
            string file = "";
            string currentKat = currentPage();
            string currentTab = readCurrentTab();

            if (currentTab == "Budget")
            {
                if (currentKat == "MDA")
                    file = "\"" + settingsPath + "\\budsjettMda.html\" ";
                else if (currentKat == "AudioVideo")
                    file = "\"" + settingsPath + "\\budsjettAudioVideo.html\" ";
                else if (currentKat == "SDA")
                    file = "\"" + settingsPath + "\\budsjettSda.html\" ";
                else if (currentKat == "Tele")
                    file = "\"" + settingsPath + "\\budsjettTele.html\" ";
                else if (currentKat == "Data")
                    file = "\"" + settingsPath + "\\budsjettData.html\" ";
                else if (currentKat == "Cross")
                    file = "\"" + settingsPath + "\\budsjettCross.html\" ";
                else if (currentKat == "Kasse")
                    file = "\"" + settingsPath + "\\budsjettKasse.html\" ";
                else if (currentKat == "Aftersales")
                    file = "\"" + settingsPath + "\\budsjettAftersales.html\" ";
                else if (currentKat == "MDASDA")
                    file = "\"" + settingsPath + "\\budsjettMdaSda.html\" ";
                else if (currentKat == "Butikk")
                    file = "\"" + settingsPath + "\\budsjettButikk.html\" ";
                else if (currentKat == "Daglig")
                    file = "\"" + settingsPath + "\\budsjettDaglig.html\" ";
            }
            else
            {
                if (currentKat == "Data")
                    file = "\"" + settingsPath + "\\rankingData.html\" ";
                else if (currentKat == "AudioVideo")
                    file = "\"" + settingsPath + "\\rankingAudioVideo.html\" ";
                else if (currentKat == "Tele")
                    file = "\"" + settingsPath + "\\rankingTele.html\" ";
                else if (currentKat == "Butikk")
                    file = "\"" + settingsPath + "\\rankingButikk.html\" ";
                else if (currentKat == "KnowHow")
                    file = "\"" + settingsPath + "\\rankingKnowHow.html\" ";
                else if (currentKat == "Oversikt")
                    file = "\"" + settingsPath + "\\rankingOversikt.html\" ";
                else if (currentKat == "Toppselgere")
                    file = "\"" + settingsPath + "\\rankingToppselgere.html\" ";
                else if (currentKat == "Lister")
                    file = "\"" + settingsPath + "\\rankingLister.html\" ";
                else if (currentKat == "Vinnprodukter")
                    file = "\"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                else if (currentKat == "VinnSelger")
                    file = "\"" + settingsPath + "\\rankingVinnprodukterSelger.html\" ";
                else if (currentKat == "Rapport" && File.Exists(htmlRapport))
                    file = "\"" + settingsPath + "\\rankingRapport.html\" ";
                else if (currentKat == "Quick" && File.Exists(htmlRankingQuick))
                    file = "\"" + settingsPath + "\\rankingQuick.html\" ";
                else if (currentKat == "Graf" && File.Exists(htmlGraf))
                    file = "\"" + settingsPath + "\\rankingGraf.html\" ";
                else if (currentKat == "Periode" && File.Exists(htmlPeriode))
                    file = "\"" + settingsPath + "\\rankingPeriode.html\" ";
                else if (currentKat == "ServiceOversikt" && File.Exists(htmlServiceOversikt))
                    file += " \"" + htmlServiceOversikt + "\" ";
                else if (currentKat == "ServiceList" && File.Exists(htmlServiceList))
                    file += " \"" + htmlServiceList + "\" ";
                else if (currentKat == "ServiceDetails" && File.Exists(htmlServiceDetails))
                    file += " \"" + htmlServiceDetails + "\" ";
                else if (currentKat == "Obsolete" && File.Exists(htmlStoreObsolete))
                    file += " \"" + htmlStoreObsolete + "\" ";
                else if (currentKat == "ObsoleteList" && File.Exists(htmlStoreObsoleteList))
                    file += " \"" + htmlStoreObsoleteList + "\" ";
                else if (currentKat.Equals("LagerUkeAnnonser") && File.Exists(htmlStoreWeekly))
                    file += " \"" + htmlStoreWeekly + "\" ";
                else if (currentKat.Equals("LagerPrisguide") && File.Exists(htmlStorePrisguide))
                    file += " \"" + htmlStorePrisguide + "\" ";
                else if (currentKat.Equals("LagerUkeAnnonserOversikt") && File.Exists(htmlStoreWeeklyOverview))
                    file += " \"" + htmlStoreWeeklyOverview + "\" ";
                else if (currentKat.Equals("LagerPrisguideOversikt") && File.Exists(htmlStorePrisguideOverview))
                    file += " \"" + htmlStorePrisguideOverview + "\" ";
                else if (currentKat == "ObsoleteImports" && File.Exists(htmlStoreObsoleteImports))
                    file += " \"" + htmlStoreObsoleteImports + "\" ";
                else if (currentKat == "Tjenester")
                    file = "\"" + settingsPath + "\\rankingAvdTjenester.html\" ";
                else if (currentKat == "Snittpriser")
                    file = "\"" + settingsPath + "\\rankingAvdSnittpriser.html\" ";
                else
                    file += " \"" + htmlImport + "\" ";
            }
            return file;
        }

        private string currentBudgetPageFile()
        {
            string file = "";
            BudgetCategory cat = currentBudgetPage();

            if (cat == BudgetCategory.MDA)
                file = "\"" + settingsPath + "\\budsjettMda.html\" ";
            else if (cat == BudgetCategory.AudioVideo)
                file = "\"" + settingsPath + "\\budsjettAudioVideo.html\" ";
            else if (cat == BudgetCategory.SDA)
                file = "\"" + settingsPath + "\\budsjettSda.html\" ";
            else if (cat == BudgetCategory.Tele)
                file = "\"" + settingsPath + "\\budsjettTele.html\" ";
            else if (cat == BudgetCategory.Data)
                file = "\"" + settingsPath + "\\budsjettData.html\" ";
            else if (cat == BudgetCategory.Cross)
                file = "\"" + settingsPath + "\\budsjettCross.html\" ";
            else if (cat == BudgetCategory.Kasse)
                file = "\"" + settingsPath + "\\budsjettKasse.html\" ";
            else if (cat == BudgetCategory.Aftersales)
                file = "\"" + settingsPath + "\\budsjettAftersales.html\" ";
            else if (cat == BudgetCategory.MDASDA)
                file = "\"" + settingsPath + "\\budsjettMdaSda.html\" ";
            else if (cat == BudgetCategory.Daglig)
                file = "\"" + settingsPath + "\\budsjettDaglig.html\" ";

            return file;
        }

        private void HighlightButton(string katArg = "")
        {
            try
            {
                buttonToppselgere.BackColor = SystemColors.ControlLight;
                buttonOversikt.BackColor = SystemColors.ControlLight;
                buttonButikk.BackColor = SystemColors.ControlLight;
                buttonData.BackColor = SystemColors.ControlLight;
                buttonAV.BackColor = SystemColors.ControlLight;
                buttonTele.BackColor = SystemColors.ControlLight;
                buttonKnowHow.BackColor = SystemColors.ControlLight;
                buttonLister.BackColor = SystemColors.ControlLight;
                buttonRankVinnprodukter.BackColor = SystemColors.ControlLight;
                buttonAvdTjenester.BackColor = SystemColors.ControlLight;
                buttonAvdSnittpriser.BackColor = SystemColors.ControlLight;

                if (katArg == "Toppselgere")
                    buttonToppselgere.BackColor = Color.LightSkyBlue;
                else if (katArg == "Oversikt")
                    buttonOversikt.BackColor = Color.LightSkyBlue;
                else if (katArg == "Butikk")
                    buttonButikk.BackColor = Color.LightSkyBlue;
                else if (katArg == "KnowHow")
                    buttonKnowHow.BackColor = Color.LightSkyBlue;
                else if (katArg == "Data")
                    buttonData.BackColor = Color.LightSkyBlue;
                else if (katArg == "AudioVideo")
                    buttonAV.BackColor = Color.LightSkyBlue;
                else if (katArg == "Tele")
                    buttonTele.BackColor = Color.LightSkyBlue;
                else if (katArg == "Lister")
                    buttonLister.BackColor = Color.LightSkyBlue;
                else if (katArg == "Vinnprodukter")
                    buttonRankVinnprodukter.BackColor = Color.LightSkyBlue;
                else if (katArg == "Tjenester")
                    buttonAvdTjenester.BackColor = Color.LightSkyBlue;
                else if (katArg == "Snittpriser")
                    buttonAvdSnittpriser.BackColor = Color.LightSkyBlue;
                this.Update();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void HighlightBudgetButton(BudgetCategory cat)
        {
            try
            {
                buttonBudgetMda.Enabled = appConfig.budgetShowMda;
                buttonBudgetAv.Enabled = appConfig.budgetShowAudioVideo;
                buttonBudgetSda.Enabled = appConfig.budgetShowSda;
                buttonBudgetTele.Enabled = appConfig.budgetShowTele;
                buttonBudgetData.Enabled = appConfig.budgetShowData;
                buttonBudgetCross.Enabled = appConfig.budgetShowCross;
                buttonBudgetKasse.Enabled = appConfig.budgetShowKasse;
                buttonBudgetAftersales.Enabled = appConfig.budgetShowAftersales;
                buttonBudgetMdaSda.Enabled = appConfig.budgetShowMdasda;
                buttonBudgetButikk.Enabled = appConfig.budgetShowButikk;

                buttonBudgetDaily.BackColor = SystemColors.ControlLight;
                buttonBudgetMda.BackColor = SystemColors.ControlLight;
                buttonBudgetAv.BackColor = SystemColors.ControlLight;
                buttonBudgetSda.BackColor = SystemColors.ControlLight;
                buttonBudgetTele.BackColor = SystemColors.ControlLight;
                buttonBudgetData.BackColor = SystemColors.ControlLight;
                buttonBudgetCross.BackColor = SystemColors.ControlLight;
                buttonBudgetKasse.BackColor = SystemColors.ControlLight;
                buttonBudgetAftersales.BackColor = SystemColors.ControlLight;
                buttonBudgetMdaSda.BackColor = SystemColors.ControlLight;
                buttonBudgetButikk.BackColor = SystemColors.ControlLight;

                if (cat == BudgetCategory.MDA)
                    buttonBudgetMda.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.AudioVideo)
                    buttonBudgetAv.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.SDA)
                    buttonBudgetSda.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Tele)
                    buttonBudgetTele.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Data)
                    buttonBudgetData.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Cross)
                    buttonBudgetCross.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Kasse)
                    buttonBudgetKasse.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Aftersales)
                    buttonBudgetAftersales.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.MDASDA)
                    buttonBudgetMdaSda.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Butikk)
                    buttonBudgetButikk.BackColor = Color.LightSkyBlue;
                else if (cat == BudgetCategory.Daglig)
                    buttonBudgetDaily.BackColor = Color.LightSkyBlue;

                this.Update();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void HighlightServiceButton(string katArg = "")
        {
            try
            {
                button27.BackColor = SystemColors.ControlLight;
                button25.BackColor = SystemColors.ControlLight;
                if (katArg == "ServiceOversikt")
                    button27.BackColor = Color.LightSkyBlue;
                else if (katArg == "ServiceList")
                    button25.BackColor = Color.LightSkyBlue;
                this.Update();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        /// <summary>
        /// Fjerner hash verdier (tvinger oppdatering av ranking sider) og lyser opp Oppdater knappen
        /// </summary>
        /// <param name="katArg">Slett spesifikk side</param>
        /// <param name="bg">Kjør bare i bakgrunn (ikke endre UI)</param>
        public void ClearHash(string katArg = "", bool bg = false)
        {
            if (!String.IsNullOrEmpty(katArg))
            {
                if (katArg == "Toppselgere")
                    appConfig.strToppselgere = "";
                else if (katArg == "Oversikt")
                    appConfig.strOversikt = "";
                else if (katArg == "Butikk")
                    appConfig.strButikk = "";
                else if (katArg == "KnowHow")
                    appConfig.strKnowHow = "";
                else if (katArg == "Data")
                    appConfig.strData = "";
                else if (katArg == "AudioVideo")
                    appConfig.strAudioVideo = "";
                else if (katArg == "Tele")
                    appConfig.strTele = "";
                else if (katArg == "ServiceOversikt")
                    appConfig.strServiceOversikt = "";
                else if (katArg == "Lister")
                    appConfig.strLister = "";
                else if (katArg == "Vinnprodukter")
                    appConfig.strVinnprodukter = "";
                else if (katArg == "Tjenester")
                    appConfig.strTjenester = "";
                else if (katArg == "Snittpriser")
                    appConfig.strSnittpriser = "";
                else if (katArg == "Obsolete")
                    appConfig.strObsolete = "";
                else if (katArg == "ObsoleteList")
                    appConfig.strObsoleteList = "";
                else if (katArg == "LagerUkeAnnonser")
                    appConfig.strLagerWeekly = "";
                else if (katArg == "LagerPrisguide")
                    appConfig.strLagerPrisguide = "";
                else if (katArg == "ObsoleteImports")
                    appConfig.strObsoleteImports = "";
            }
            else
            {
                appConfig.strToppselgere = "";
                appConfig.strOversikt = "";
                appConfig.strButikk = "";
                appConfig.strKnowHow = "";
                appConfig.strData = "";
                appConfig.strAudioVideo = "";
                appConfig.strTele = "";
                appConfig.strServiceOversikt = "";
                appConfig.strLister = "";
                appConfig.strVinnprodukter = "";
                appConfig.strTjenester = "";
                appConfig.strSnittpriser = "";
                appConfig.strObsolete = "";
                appConfig.strObsoleteList = "";
                appConfig.strLagerWeekly = "";
                appConfig.strLagerPrisguide = "";
                appConfig.strObsoleteImports = "";
            }
            SaveSettings();
            if (!bg && !EmptyDatabase())
            {
                buttonOppdater.BackColor = Color.LightGreen;
                buttonOppdater.ForeColor = SystemColors.ControlText;
            }
            if (!bg && service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
            {
                buttonServiceOppdater.BackColor = Color.LightGreen;
                buttonServiceOppdater.ForeColor = SystemColors.ControlText;
            }
        }

        private void ClearBudgetHash(BudgetCategory cat, bool bg = false)
        {
            if (cat != BudgetCategory.None)
            {
                if (cat == BudgetCategory.MDA)
                    appConfig.strBudgetMda = "";
                else if (cat == BudgetCategory.AudioVideo)
                    appConfig.strBudgetAv = "";
                else if (cat == BudgetCategory.SDA)
                    appConfig.strBudgetSda = "";
                else if (cat == BudgetCategory.Tele)
                    appConfig.strBudgetTele = "";
                else if (cat == BudgetCategory.Data)
                    appConfig.strBudgetData = "";
                else if (cat == BudgetCategory.Cross)
                    appConfig.strBudgetCross = "";
                else if (cat == BudgetCategory.Kasse)
                    appConfig.strBudgetKasse = "";
                else if (cat == BudgetCategory.Aftersales)
                    appConfig.strBudgetAftersales = "";
                else if (cat == BudgetCategory.MDASDA)
                    appConfig.strBudgetMdasda = "";
                else if (cat == BudgetCategory.Daglig)
                    appConfig.strBudgetDaily = "";
            }
            else
            {
                appConfig.strBudgetMda = "";
                appConfig.strBudgetAv = "";
                appConfig.strBudgetSda = "";
                appConfig.strBudgetTele = "";
                appConfig.strBudgetData = "";
                appConfig.strBudgetCross = "";
                appConfig.strBudgetKasse = "";
                appConfig.strBudgetAftersales = "";
                appConfig.strBudgetMdasda = "";
                appConfig.strBudgetDaily = "";
            }
            SaveSettings();
            if (!bg && !EmptyDatabase())
            {
                buttonBudgetUpdate.BackColor = Color.LightGreen;
                buttonBudgetUpdate.ForeColor = SystemColors.ControlText;
            }
        }

        /// <summary>
        /// Sjekk om programmet er opptatt med andre oppgaver.
        /// </summary>
        /// <param name="checkAll">Hvis true, sjekker også ranking prosesser.</param>
        /// <returns>True = vi er opptatt, False = ingen operasjoner igang</returns>
        [MethodImplAttribute(MethodImplOptions.NoInlining)]
        public bool IsBusy(bool checkAll = false, bool silent = false)
        {
            if (!bwAutoRanking.IsBusy && !bwPDF.IsBusy && !bwImport.IsBusy && !bwAutoImportService.IsBusy &&
                !bwImportService.IsBusy && !bwImportObsolete.IsBusy && !bwQuickAuto.IsBusy && !bwAutoStore.IsBusy &&
                !bwMacroRanking.IsBusy && !appManagerIsBusy && !worker.IsBusy && !checkAll)
                return false;

            if (!bwGraph.IsBusy && !bwMakeScreens.IsBusy && !bwQuickAuto.IsBusy && !bwService.IsBusy &&
                !bwServiceReport.IsBusy && !bwPopulateSk.IsBusy && !bwUpdateBigGraph.IsBusy && !bwAutoImportService.IsBusy &&
                !bwImportService.IsBusy && !bwRanking.IsBusy && !bwStore.IsBusy && !bwReport.IsBusy && !bwSendEmail.IsBusy &&
                !bwImport.IsBusy && !bwPDF.IsBusy && !bwAutoRanking.IsBusy && !bwAutoStore.IsBusy && !bwImportObsolete.IsBusy &&
                !bwMacroRanking.IsBusy && !bwBudget.IsBusy && !appManagerIsBusy && !worker.IsBusy && checkAll)
                return false;

            if (Loaded && !silent)
                Logg.Log("Vent litt, er opptatt med andre oppgaver..", Color.Black);

            if (appConfig.debug)
            {
                string strBusy = "IsBusy(" + checkAll + ") [" + new StackFrame(1).GetMethod().Name + "] prosess(er) opptatt: ";
                if (bwGraph.IsBusy) strBusy += " bwGraph";
                if (bwMakeScreens.IsBusy) strBusy += " bwMakeScreens";
                if (bwQuickAuto.IsBusy) strBusy += " bwQuickAuto";
                if (bwService.IsBusy) strBusy += " bwService";
                if (bwServiceReport.IsBusy) strBusy += " bwServiceReport";
                if (bwPopulateSk.IsBusy) strBusy += " bwPopulateSk";
                if (bwUpdateBigGraph.IsBusy) strBusy += " bwUpdateBigGraph";
                if (bwAutoImportService.IsBusy) strBusy += " bwAutoImportService";
                if (bwImportService.IsBusy) strBusy += " bwImportService";
                if (bwRanking.IsBusy) strBusy += " bwRanking";
                if (bwStore.IsBusy) strBusy += " bwStore";
                if (bwReport.IsBusy) strBusy += " bwReport";
                if (bwSendEmail.IsBusy) strBusy += " bwSendEmail";
                if (bwImport.IsBusy) strBusy += " bwImport";
                if (bwPDF.IsBusy) strBusy += " bwPDF";
                if (bwAutoRanking.IsBusy) strBusy += " bwAutoRanking";
                if (bwAutoStore.IsBusy) strBusy += " bwAutoStore";
                if (bwImportObsolete.IsBusy) strBusy += " bwImportObsolete";
                if (bwMacroRanking.IsBusy) strBusy += " bwMacroRanking";
                if (bwBudget.IsBusy) strBusy += " bwBudget";
                if (appManagerIsBusy) strBusy += " appManagerIsBusy";
                if (worker.IsBusy) strBusy += " worker";
                Logg.Debug(strBusy);
            }

            return true;
        }

        public bool EmptyDatabase()
        {
            try
            {
                if (appConfig.dbFrom.Date.Equals(DateTime.Now.Date))
                    return true;
                if (appConfig.dbFrom == DateTime.MinValue)
                    return true;
                if (appConfig.dbFrom > appConfig.dbTo)
                    return true;
                if (appConfig.dbFrom.Date == rangeMin.Date)
                    return true;
                else
                    return false; // OK det ser ut som databasen er IKKE tom.
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return true;
            }
        }

        private void loadLog()
        {
            try
            {
                richLog.Text = File.ReadAllText(settingsPath + @"\Log.txt", Encoding.Unicode);
                richLog.ScrollToCaret();
                Logg.Log("Åpnet log.");
            }
            catch(Exception ex)
            {
                FormError errorMsg = new FormError("Feill ved lasting av logg.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void moveDate(int m = 0, bool reload = false)
        {
            try
            {
                if (!EmptyDatabase())
                {
                    var d = pickerRankingDate.Value;
                    if (m == 1) // gå tilbake en måned
                    {
                        if (appConfig.dbFrom.Date <= d.AddMonths(-1))
                            pickerRankingDate.Value = d.AddMonths(-1);
                        else
                            pickerRankingDate.Value = appConfig.dbFrom;
                    }
                    if (m == 2) // gå tilbake en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbFrom.Date <= d.AddDays(-1) && d.AddDays(-1).DayOfWeek != DayOfWeek.Sunday)
                                pickerRankingDate.Value = d.AddDays(-1);
                            if (appConfig.dbFrom.Date <= d.AddDays(-2) && d.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
                                pickerRankingDate.Value = d.AddDays(-2);
                        }
                        else
                        {
                            if (appConfig.dbFrom.Date <= d.AddDays(-1))
                                pickerRankingDate.Value = d.AddDays(-1);
                        }
                    }
                    if (m == 3) // gå fram en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbTo.Date >= d.AddDays(1) && d.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                                pickerRankingDate.Value = d.AddDays(1);
                            if (appConfig.dbTo.Date >= d.AddDays(2) && d.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                                pickerRankingDate.Value = d.AddDays(2);
                        }
                        else
                        {
                            if (appConfig.dbTo.Date >= d.AddDays(1))
                                pickerRankingDate.Value = d.AddDays(1);
                        }
                    }
                    if (m == 4) // gå fram en måned
                    {
                        if (appConfig.dbTo.Date >= d.AddMonths(1))
                            pickerRankingDate.Value = d.AddMonths(1);
                        else
                            pickerRankingDate.Value = appConfig.dbTo;
                    }
                    d = pickerRankingDate.Value;
                    if (d.Date >= appConfig.dbTo.Date)
                    {
                        buttonRankF.Enabled = false; // fremover knapp
                        buttonRankFF.Enabled = false; // fremover knapp
                    }
                    else
                    {
                        buttonRankF.Enabled = true; // fremover knapp
                        buttonRankFF.Enabled = true; // fremover knapp
                    }
                    if (d.Date <= appConfig.dbFrom.Date)
                    {
                        buttonRankBF.Enabled = false; // bakover knapp
                        buttonRankB.Enabled = false; // bakover knapp
                    }
                    else
                    {
                        buttonRankBF.Enabled = true; // bakover knapp
                        buttonRankB.Enabled = true; // bakover knapp
                    }

                    if (Loaded && reload)
                    {
                        if (!IsBusy())
                            UpdateRank();
                        highlightDate = pickerRankingDate.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void moveBudgetDate(int m = 0, bool reload = false)
        {
            try
            {
                if (!EmptyDatabase())
                {
                    var d = pickerBudget.Value;
                    if (m == 1) // gå tilbake en måned
                    {
                        if (appConfig.dbFrom.Date <= d.AddMonths(-1))
                            pickerBudget.Value = d.AddMonths(-1);
                        else
                            pickerBudget.Value = appConfig.dbFrom;
                    }

                    if (m == 2) // gå tilbake en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbFrom.Date <= d.AddDays(-1) && d.AddDays(-1).DayOfWeek != DayOfWeek.Sunday)
                                pickerBudget.Value = d.AddDays(-1);
                            if (appConfig.dbFrom.Date <= d.AddDays(-2) && d.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
                                pickerBudget.Value = d.AddDays(-2);
                        }
                        else
                        {
                            if (appConfig.dbFrom.Date <= d.AddDays(-1))
                                pickerBudget.Value = d.AddDays(-1);
                        }
                    }
                    if (m == 3) // gå fram en dag
                    {
                        if (appConfig.ignoreSunday)
                        {
                            if (appConfig.dbTo.Date >= d.AddDays(1) && d.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                                pickerBudget.Value = d.AddDays(1);
                            if (appConfig.dbTo.Date >= d.AddDays(2) && d.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                                pickerBudget.Value = d.AddDays(2);
                        }
                        else
                        {
                            if (appConfig.dbTo.Date >= d.AddDays(1))
                                pickerBudget.Value = d.AddDays(1);
                        }
                    }

                    if (m == 4) // gå fram en måned
                    {
                        if (appConfig.dbTo.Date >= d.AddMonths(1))
                            pickerBudget.Value = d.AddMonths(1);
                        else
                            pickerBudget.Value = appConfig.dbTo;
                    }
                    d = pickerBudget.Value;
                    if (d.Date >= appConfig.dbTo.Date)
                    {
                        buttonBudgetFF.Enabled = false; // fremover knapp
                        buttonBudgetF.Enabled = false;
                    }
                    else
                    {
                        buttonBudgetFF.Enabled = true; // fremover knapp
                        buttonBudgetF.Enabled = true;
                    }
                    if (d.Date <= appConfig.dbFrom.Date)
                    {
                        buttonBudgetBF.Enabled = false; // bakover knapp
                        buttonBudgetB.Enabled = false;
                    }
                    else
                    {
                        buttonBudgetBF.Enabled = true; // bakover knapp
                        buttonBudgetB.Enabled = true;
                    }

                    if (Loaded && reload)
                    {
                        if (!IsBusy())
                            UpdateBudget();
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public string CreatePDF(string type = "", string file = "", BackgroundWorker bw = null, bool fileOverride = false)
        {
            try
            {
                string datoStr = (autoMode) ? datoStr = appConfig.dbTo.ToString("yyy-MM-dd") : pickerRankingDate.Value.ToString("yyy-MM-dd");
                string sourceFiles = "";
                string destinationFile = "";
                string graphStr1 = settingsPath + @"\rankingGrafOversikt.html";
                string graphStr2 = settingsPath + @"\rankingGrafButikk.html";
                string graphStr2b = settingsPath + @"\rankingGrafKnowHow.html";
                string graphStr3 = settingsPath + @"\rankingGrafData.html";
                string graphStr4 = settingsPath + @"\rankingGrafNettbrett.html";
                string graphStr5 = settingsPath + @"\rankingGrafAudioVideo.html";
                string graphStr6 = settingsPath + @"\rankingGrafTele.html";
                string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                string newHashBudget = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                string newHashStore = appConfig.Avdeling + pickerLagerDato.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                string newHashService = appConfig.Avdeling + pickerServiceDato.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();


                Logg.Log("Klargjør filer for konvertering..");
                processing.SetText = "Klargjør filer for konvertering..";

                if (!String.IsNullOrEmpty(type))
                {
                    if (type == "Vinnprodukter")
                    {
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "FullBudget")
                    {
                        if (appConfig.budgetShowMda)
                        {
                            BuildBudget(BudgetCategory.MDA, appConfig.strBudgetMda, htmlBudgetMda);
                            appConfig.strBudgetMda = newHashBudget;
                            if (File.Exists(htmlBudgetMda))
                                sourceFiles += " \"" + settingsPath + "\\budsjettMda.html\" ";
                        }
                        else if (appConfig.budgetShowAudioVideo)
                        {
                            BuildBudget(BudgetCategory.AudioVideo, appConfig.strBudgetAv, htmlBudgetAudioVideo);
                            appConfig.strBudgetAv = newHashBudget;
                            if (File.Exists(htmlBudgetAudioVideo))
                                sourceFiles += " \"" + settingsPath + "\\budsjettAudioVideo.html\" ";
                        }
                        else if (appConfig.budgetShowSda)
                        {
                            BuildBudget(BudgetCategory.SDA, appConfig.strBudgetSda, htmlBudgetSda);
                            appConfig.strBudgetSda = newHashBudget;
                            if (File.Exists(htmlBudgetSda))
                                sourceFiles += " \"" + settingsPath + "\\budsjettSda.html\" ";
                        }
                        else if (appConfig.budgetShowTele)
                        {
                            BuildBudget(BudgetCategory.Tele, appConfig.strBudgetTele, htmlBudgetTele);
                            appConfig.strBudgetTele = newHashBudget;
                            if (File.Exists(htmlBudgetTele))
                                sourceFiles += " \"" + settingsPath + "\\budsjettTele.html\" ";
                        }
                        else if (appConfig.budgetShowData)
                        {
                            BuildBudget(BudgetCategory.Data, appConfig.strBudgetData, htmlBudgetData);
                            appConfig.strBudgetData = newHashBudget;
                            if (File.Exists(htmlBudgetData))
                                sourceFiles += " \"" + settingsPath + "\\budsjettData.html\" ";
                        }
                        else if (appConfig.budgetShowCross)
                        {
                            BuildBudget(BudgetCategory.Cross, appConfig.strBudgetCross, htmlBudgetCross);
                            appConfig.strBudgetCross = newHashBudget;
                            if (File.Exists(htmlBudgetCross))
                                sourceFiles += " \"" + settingsPath + "\\budsjettCross.html\" ";
                        }
                        else if (appConfig.budgetShowKasse)
                        {
                            BuildBudget(BudgetCategory.Kasse, appConfig.strBudgetKasse, htmlBudgetKasse);
                            appConfig.strBudgetKasse = newHashBudget;
                            if (File.Exists(htmlBudgetKasse))
                                sourceFiles += " \"" + settingsPath + "\\budsjettKasse.html\" ";
                        }
                        else if (appConfig.budgetShowAftersales)
                        {
                            BuildBudget(BudgetCategory.Aftersales, appConfig.strBudgetAftersales, htmlBudgetAftersales);
                            appConfig.strBudgetAftersales = newHashBudget;
                            if (File.Exists(htmlBudgetAftersales))
                                sourceFiles += " \"" + settingsPath + "\\budsjettAftersales.html\" ";
                        }
                        else if (appConfig.budgetShowMdasda)
                        {
                            BuildBudget(BudgetCategory.MDASDA, appConfig.strBudgetMdasda, htmlBudgetMdasda);
                            appConfig.strBudgetMdasda = newHashBudget;
                            if (File.Exists(htmlBudgetMdasda))
                                sourceFiles += " \"" + settingsPath + "\\budsjettMdaSda.html\" ";
                        }
                        if (String.IsNullOrEmpty(sourceFiles))
                        {
                            Logg.Log("Ingen budsjett er klar til eksportering til PDF. Opprett nye budsjett og/eller sjekk over budsjett innstillinger.", Color.Red);
                            return "";
                        }
                    }
                    else if (type == "KnowHow")
                    {
                        if (appConfig.pdfVisKnowHow)
                        {
                            BuildKnowHowRanking(true);
                            appConfig.strKnowHow = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingKnowHow.html\" ";
                            if (appConfig.graphKnowHow && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("KnowHow", true, bw);
                                sourceFiles += " \"" + graphStr2b + "\" ";
                            }
                        }
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "Computer")
                    {
                        if (appConfig.pdfVisData)
                        {
                            BuildDataRanking(true, bw);
                            appConfig.strData = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingData.html\" ";
                            if (appConfig.graphData && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Data", true, bw);
                                sourceFiles += " \"" + graphStr3 + "\" ";
                                ViewGraph("Nettbrett", true, bw);
                                sourceFiles += " \"" + graphStr4 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "AudioVideo")
                    {
                        if (appConfig.pdfVisAudioVideo)
                        {
                            BuildAudioVideoRanking(true, bw);
                            appConfig.strAudioVideo = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingAudioVideo.html\" ";
                            if (appConfig.graphAudioVideo && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("AudioVideo", true, bw);
                                sourceFiles += " \"" + graphStr5 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "Telecom")
                    {
                        if (appConfig.pdfVisTele)
                        {
                            BuildTeleRanking(true, bw);
                            appConfig.strTele = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingTele.html\" ";
                            if (appConfig.graphTele && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Tele", true, bw);
                                sourceFiles += " \"" + graphStr6 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "Cross")
                    {
                        if (appConfig.pdfVisKnowHow)
                        {
                            BuildKnowHowRanking(true);
                            appConfig.strKnowHow = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingKnowHow.html\" ";
                            if (appConfig.graphKnowHow && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("KnowHow", true, bw);
                                sourceFiles += " \"" + graphStr2b + "\" ";
                            }
                        }
                        if (appConfig.pdfVisData)
                        {
                            BuildDataRanking(true, bw);
                            appConfig.strData = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingData.html\" ";
                            if (appConfig.graphData && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Data", true, bw);
                                sourceFiles += " \"" + graphStr3 + "\" ";
                                ViewGraph("Nettbrett", true, bw);
                                sourceFiles += " \"" + graphStr4 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisAudioVideo)
                        {
                            BuildAudioVideoRanking(true, bw);
                            appConfig.strAudioVideo = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingAudioVideo.html\" ";
                            if (appConfig.graphAudioVideo && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("AudioVideo", true, bw);
                                sourceFiles += " \"" + graphStr5 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisTele)
                        {
                            BuildTeleRanking(true, bw);
                            appConfig.strTele = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingTele.html\" ";
                            if (appConfig.graphTele && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Tele", true, bw);
                                sourceFiles += " \"" + graphStr6 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisVinnprodukter)
                        {
                            BuildVinnRanking(true);
                            appConfig.strVinnprodukter = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                        }
                    }
                    else if (type == "Lister")
                    {
                        if (appConfig.pdfVisLister)
                        {
                            BuildListerRanking(true);
                            appConfig.strLister = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingLister.html\" ";
                        }
                    }
                    else if (type == "Full" || type == "FullUkestart")
                    {
                        if (appConfig.importSetting.StartsWith("Full"))
                        {
                            if (appConfig.pdfVisToppselgere)
                            {
                                BuildToppselgereRanking(true);
                                appConfig.strToppselgere = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingToppselgere.html\" ";
                            }
                            if (appConfig.pdfVisOversikt && appConfig.importSetting.StartsWith("Full"))
                            {
                                BuildOversiktRanking(true, bw);
                                appConfig.strOversikt = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingOversikt.html\" ";
                                if (appConfig.pdfExpandedGraphs && appConfig.graphOversikt && File.Exists(graphStr1))
                                {
                                    ViewGraph("Oversikt", true, bw);
                                    sourceFiles += " \"" + graphStr1 + "\" ";
                                }
                            }
                        }
                        if (appConfig.pdfVisButikk)
                        {
                            BuildButikkRanking(true, bw);
                            appConfig.strButikk = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingButikk.html\" ";
                            if (appConfig.pdfExpandedGraphs && appConfig.graphButikk)
                            {
                                ViewGraph("Butikk", true, bw);
                                sourceFiles += " \"" + graphStr2 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisKnowHow)
                        {
                            BuildKnowHowRanking(true);
                            appConfig.strKnowHow = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingKnowHow.html\" ";
                            if (appConfig.graphKnowHow && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("KnowHow", true, bw);
                                sourceFiles += " \"" + graphStr2b + "\" ";
                            }
                        }
                        if (appConfig.pdfVisData)
                        {
                            BuildDataRanking(true, bw);
                            appConfig.strData = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingData.html\" ";
                            if (appConfig.graphData && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Data", true, bw);
                                sourceFiles += " \"" + graphStr3 + "\" ";
                                ViewGraph("Nettbrett", true, bw);
                                sourceFiles += " \"" + graphStr4 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisAudioVideo)
                        {
                            BuildAudioVideoRanking(true, bw);
                            appConfig.strAudioVideo = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingAudioVideo.html\" ";
                            if (appConfig.graphAudioVideo && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("AudioVideo", true, bw);
                                sourceFiles += " \"" + graphStr5 + "\" ";
                            }
                        }
                        if (appConfig.pdfVisTele)
                        {
                            BuildTeleRanking(true, bw);
                            appConfig.strTele = newHash;
                            sourceFiles += " \"" + settingsPath + "\\rankingTele.html\" ";
                            if (appConfig.graphTele && appConfig.pdfExpandedGraphs)
                            {
                                ViewGraph("Tele", true, bw);
                                sourceFiles += " \"" + graphStr6 + "\" ";
                            }
                        }
                        if (appConfig.importSetting.StartsWith("Full"))
                        {
                            if (appConfig.pdfVisLister)
                            {
                                BuildListerRanking(true);
                                appConfig.strLister = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingLister.html\" ";
                            }
                            if (appConfig.pdfVisVinnprodukter)
                            {
                                BuildVinnRanking(true);
                                appConfig.strVinnprodukter = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingVinnprodukter.html\" ";
                            }
                            if (appConfig.pdfVisTjenester)
                            {
                                BuildAvdTjenester(true);
                                appConfig.strTjenester = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingAvdTjenester.html\" ";
                            }
                            if (appConfig.pdfVisSnittpriser)
                            {
                                BuildAvdSnittpriser(true);
                                appConfig.strSnittpriser = newHash;
                                sourceFiles += " \"" + settingsPath + "\\rankingAvdSnittpriser.html\" ";
                            }
                        }
                        if (!EmptyStoreDatabase())
                        {
                            if (appConfig.pdfVisObsolete)
                            {
                                BuildStoreStatus(true);
                                appConfig.strObsolete = newHashStore;
                                sourceFiles += " \"" + htmlStoreObsolete + "\" ";
                            }
                            if (appConfig.pdfVisObsoleteList)
                            {
                                BuildStoreObsoleteList(true);
                                appConfig.strObsoleteList = newHashStore;
                                sourceFiles += " \"" + htmlStoreObsoleteList + "\" ";
                            }
                            if (appConfig.pdfVisWeekly)
                            {
                                BuildStoreWeeklyOverview(true);
                                appConfig.strLagerWeeklyOverview = newHashStore;
                                sourceFiles += " \"" + htmlStoreWeeklyOverview + "\" ";
                            }
                            if (appConfig.pdfVisPrisguide)
                            {
                                BuildStorePrisguideOverview(true);
                                appConfig.strLagerPrisguideOverview = newHashStore;
                                sourceFiles += " \"" + htmlStorePrisguideOverview + "\" ";
                            }
                        }
                        if (appConfig.pdfVisService && service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
                        {
                            MakeServiceOversikt(true, bw);
                            appConfig.strServiceOversikt = newHashService;
                            sourceFiles += " \"" + htmlServiceOversikt + "\" ";
                        }
                    }
                    destinationFile = "\"" + System.IO.Path.GetTempPath() + "TjenesteRanking " + type + " " + appConfig.Avdeling + " " + datoStr + ".pdf\"";
                }
                else if (!String.IsNullOrEmpty(file))
                {
                    sourceFiles = file;
                    destinationFile = "\"" + settingsTemp + "\\KGSA " + appConfig.Avdeling + " " + currentPage() + " " + datoStr + ".pdf\"";
                    //destinationFile = "\"" + System.IO.Path.GetTempPath() + "KGSA " + appConfig.Avdeling + " " + currentPage() + " " + datoStr + ".pdf\"";
                }

                if (destinationFile != null)
                    filenamePDF = destinationFile.Substring(1, destinationFile.Length - 2);
                else
                {
                    Logg.Log("Feil oppstod under generering av PDF: Ingen destinasjon valgt.", Color.Red);
                    return "";
                }

                if (String.IsNullOrEmpty(sourceFiles))
                {
                    Logg.Log("Error: Har ikke data fra alle kategoriene til å lage en fullstendig rapport.", Color.Red);
                    return ""; // mangler filer
                }

                Logg.Log("Konverterer til PDF..");
                processing.SetText = "Konvertere til PDF..";

                if (fileOverride && filenamePDF.Length > 3)
                {
                    try
                    {
                        if (File.Exists(filenamePDF))
                            File.Delete(filenamePDF);
                    }
                    catch
                    {
                        Logg.Debug("PDF (" + filenamePDF + ") er låst. Endrer filnavn..");
                        if (destinationFile.Length > 3)
                        {
                            Random r = new Random();
                            destinationFile = destinationFile.Substring(0, filenamePDF.Length - 3) + " " + r.Next(99) + ".pdf\"";
                            filenamePDF = destinationFile.Substring(1, destinationFile.Length - 2);
                        }
                    }
                }

                string options = " -B 15 -L 7 -R 7 -T 7 --zoom " + appConfig.pdfZoom + " ";
                if (type == "Full" && appConfig.pdfTableOfContent)
                    options = " --title \"KGSA Tjenesteranking " + datoStr + "\" -B 20 -L 7 -R 7 -T 7 --zoom " + appConfig.pdfZoom + " toc --toc-header-text \"KGSA Tjenesteranking " + datoStr + "\" --toc-header-text \" Innhold - KGSA Tjenesteranking\" ";

                if (appConfig.pdfLandscape)
                    options += "-O landscape ";

                Logg.Debug("PDF generator: wkhtmltopdf " + options + sourceFiles + destinationFile);

                var wkhtmltopdf = new ProcessStartInfo();
                wkhtmltopdf.WindowStyle = ProcessWindowStyle.Hidden;
                wkhtmltopdf.FileName = filePDFwkhtmltopdf;
                wkhtmltopdf.Arguments = options + sourceFiles + destinationFile;
                wkhtmltopdf.WorkingDirectory = settingsPath;
                wkhtmltopdf.CreateNoWindow = true;
                wkhtmltopdf.UseShellExecute = false;

                Process D = Process.Start(wkhtmltopdf);

                D.WaitForExit(20000);

                if (!D.HasExited)
                {
                    Logg.Log("Error: PDF generatoren ble tidsavbrutt.", Color.Red);
                    return "";
                }

                int result = D.ExitCode;
                if (result != 0)
                {
                    Logg.Log("Error: PDF generator returnerte med feilkode " + result, Color.Red);
                    return "";
                }
                
                return filenamePDF;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        public void LoadSettings()
        {
            if (!File.Exists(settingsFile))
            {
                appConfig = new AppSettings();
                SaveSettings();
            }

            try
            {
                XmlSerializer mySerializer = new XmlSerializer(typeof(AppSettings));
                using (StreamReader myXmlReader = new StreamReader(settingsFile))
                {
                    appConfig = (AppSettings)mySerializer.Deserialize(myXmlReader);

                    try
                    {
                        // SJEKK LEGACY INNSTILLINGER  -----   START
                        if (!appConfig.storeObsoleteSortBy.Contains("tblUkurans."))
                            appConfig.storeObsoleteSortBy = "tblUkurans.UkuransProsent";

                        if (String.IsNullOrEmpty(appConfig.visningNull))
                            appConfig.visningNull = "&nbsp;"; // visningNull som "" ødelegger CSS styles i tabellen, sett som &nbsp;

                        if (appConfig.varekoder == null)
                        {
                            appConfig.varekoder = new List<VarekodeList> { };
                            appConfig.varekoder.Clear();
                            appConfig.ResetVarekoder(appConfig.varekoder);
                        }

                        if (appConfig.varekoder.Count == 0)
                        {
                            appConfig.varekoder = new List<VarekodeList> { };
                            appConfig.varekoder.Clear();
                            appConfig.ResetVarekoder(appConfig.varekoder);
                        }

                        if (appConfig.graphResX < 711 && appConfig.graphResY < 251) // Automatisk øk oppløsningen hvis settings inneholder for lav gammel oppløsning
                        {
                            appConfig.graphResX = 900;
                            appConfig.graphResY = 300;
                        }

                        if (newInstall && macroProgramQuick != null && macroProgram != null) // Se etter utdaterte makroer og anbefal reset..
                            if (File.Exists(macroProgramQuick) && File.Exists(macroProgram))
                                if (!File.ReadAllText(macroProgramQuick).Contains("ImportKveldstall()") && !File.ReadAllText(macroProgram).Contains("Start("))
                                {
                                    Logg.Alert("Bare en liten ting før vi fortsetter..\n\nMakro programmene som benyttes i automatisk importering av transaksjoner og ranking var utdatert og er blitt satt tilbake til nytt standard format.\nObs! Nødvendige endringer må gjøres før de kan benyttes igjen.", "KGSA - Makro program oppdatert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    Logg.Log("Resetter makroer..");
                                    File.Delete(macroProgramQuick);
                                    File.Delete(macroProgramService);
                                    File.Delete(macroProgram);
                                    StartupCheck();
                                    Logg.Log("Makroer endret til standard. Må redigeres før bruk!");
                                }

                        // SJEKK LEGACY INNSTILLINGER  -----   STOP
                    }
                    catch (Exception ex)
                    {
                        Logg.Log("Kritisk feil under lasting av innstillinger: " + ex.Message);
                        Logg.Unhandled(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Log("Feil ved lesing av settings.xnl. Forsøker å ta kopi av fil før gjenoppretting..", Color.Red, true);
                try
                {
                    if (File.Exists(FormMain.settingsPath + @"\settings.xml.bk"))
                    {
                        Logg.Debug("Sletter backup settings.xml.bk..");
                        File.Delete(FormMain.settingsPath + @"\settings.xml.bk");
                    }
                    Logg.Debug("Kopierer " + FormMain.settingsFile + " til settings.xml.bk..");
                    File.Copy(FormMain.settingsFile, FormMain.settingsPath + @"\settings.xml.bk");
                    Logg.Debug("Sletter slettings settingx.xml..");
                    File.Delete(FormMain.settingsFile);
                    appConfig = new AppSettings();
                    ReloadDatabase();
                    SaveSettings();
                }
                catch (Exception ix)
                {
                    Logg.Unhandled(ix);
                }
                Logg.Unhandled(ex);
                Logg.Alert("Feil ved lasting av innstillinger\n\nStandard innstillinger vil bli lastet inn etter en omstart av programmet.\n\nEn kopi av de gamle innstillingene ble kopiert til: " + FormMain.settingsPath + @"\settings.xml.bk", "Innstillinger tilbakestilt", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                forceShutdown = true;
                this.Close();
            }
        }

        public void SaveSettings()
        {
            try
            { 
                var serializerObj = new XmlSerializer(typeof(AppSettings));
                TextWriter writeFileStream = new StreamWriter(settingsFile);
                using (writeFileStream)
                {
                    serializerObj.Serialize(writeFileStream, appConfig);
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Uhåndtert feil under lagring av innstillinger.", ex);
                errorMsg.ShowDialog();
                return;
            }
        }

        public static void StartupCheck()
        {
            try
            {
                try
                {
                    if (!Directory.Exists(System.IO.Path.GetTempPath()))
                    {
                        Directory.CreateDirectory(System.IO.Path.GetTempPath());
                    }
                }
                catch(Exception ex)
                {
                    Logg.Unhandled(ex);
                }
                if (!Directory.Exists(settingsPath))
                {
                    Directory.CreateDirectory(settingsPath);
                }
                if (!Directory.Exists(settingsTemp))
                {
                    Directory.CreateDirectory(settingsTemp);
                }
                try {
                    var files = new DirectoryInfo(settingsTemp).GetFiles("*.*");
                    foreach (var file in files.Where(file => DateTime.UtcNow - file.CreationTimeUtc > TimeSpan.FromHours(72)))
                    {
                        file.Delete();
                    }
                }
                catch
                { // do nothing
                }
                if (File.Exists(settingsPath + @"\install.txt"))
                {
                    newInstall = true;
                    File.Delete(settingsPath + @"\install.txt");
                }
                if (!File.Exists(htmlImport))
                {
                    File.WriteAllText(htmlImport, Resources.htmlImportData, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlImport).Length != Resources.htmlImportData.Length)
                {
                    File.Delete(htmlImport);
                    File.WriteAllText(htmlImport, Resources.htmlImportData, Encoding.Unicode);
                }
                if (!File.Exists(htmlImportStore))
                {
                    File.WriteAllText(htmlImportStore, Resources.htmlImportStore, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlImportStore).Length != Resources.htmlImportStore.Length)
                {
                    File.Delete(htmlImportStore);
                    File.WriteAllText(htmlImportStore, Resources.htmlImportStore, Encoding.Unicode);
                }
                if (!File.Exists(htmlSetupBudget))
                {
                    File.WriteAllText(htmlSetupBudget, Resources.htmlSetupBudget, Encoding.Unicode);
                }
                if (!File.Exists(htmlImportService))
                {
                    File.WriteAllText(htmlImportService, Resources.htmlImportService, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlImportService).Length != Resources.htmlImportService.Length)
                {
                    File.Delete(htmlImportService);
                    File.WriteAllText(htmlImportService, Resources.htmlImportService, Encoding.Unicode);
                }
                if (!File.Exists(htmlLoading))
                {
                    File.WriteAllText(htmlLoading, Resources.htmlLoading, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlLoading).Length != Resources.htmlLoading.Length)
                {
                    File.Delete(htmlLoading);
                    File.WriteAllText(htmlLoading, Resources.htmlLoading, Encoding.Unicode);
                }

                if (!File.Exists(htmlError))
                {
                    File.WriteAllText(htmlError, Resources.htmlError, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlError).Length != Resources.htmlError.Length)
                {
                    File.Delete(htmlError);
                    File.WriteAllText(htmlError, Resources.htmlError, Encoding.Unicode);
                }
                if (!File.Exists(htmlStopped))
                {
                    File.WriteAllText(htmlStopped, Resources.htmlStopped, Encoding.Unicode);
                }
                else if (File.ReadAllText(htmlStopped).Length != Resources.htmlStopped.Length)
                {
                    File.Delete(htmlStopped);
                    File.WriteAllText(htmlStopped, Resources.htmlStopped, Encoding.Unicode);
                }
                if (!File.Exists(settingsPath + @"\favicon.ico"))
                {
                    File.WriteAllBytes(settingsPath + @"\favicon.ico", Resources.favicon);
                }
                if (!File.Exists(settingsPath + @"\bar_green.png"))
                {
                    File.WriteAllBytes(settingsPath + @"\bar_green.png", Resources.img_bar_green);
                }
                if (!File.Exists(settingsPath + @"\bar_red.png"))
                {
                    File.WriteAllBytes(settingsPath + @"\bar_red.png", Resources.img_bar_red);
                }
                if (!File.Exists(macroProgram))
                {
                    File.WriteAllText(macroProgram, Resources.Macro);
                }
                if (!File.Exists(macroProgramQuick))
                {
                    File.WriteAllText(macroProgramQuick, Resources.MacroQuick);
                }
                if (!File.Exists(macroProgramService))
                {
                    File.WriteAllText(macroProgramService, Resources.MacroService);
                }
                if (!File.Exists(macroProgramStore))
                {
                    File.WriteAllText(macroProgramStore, Resources.MacroStore);
                }
                if (!File.Exists(jsJqueryTablesorter))
                {
                    File.WriteAllText(jsJqueryTablesorter, Resources.jquery_tablesorter);
                }
                if (!File.Exists(jsJqueryMetadata))
                {
                    File.WriteAllText(jsJqueryMetadata, Resources.jquery_metadata);
                }
                if (!File.Exists(jsJquery))
                {
                    File.WriteAllText(jsJquery, Resources.jquery);
                }
                else if (File.ReadAllText(jsJquery).Length != Resources.jquery.Length)
                {
                    File.Delete(jsJquery);
                    File.WriteAllText(jsJquery, Resources.jquery);
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
            
        }

        private void NewInstallation()
        {
            try
            {
                Logg.Log("Oppstart: Ny installasjon. Setter standard innstillinger..");

                File.Delete(settingsFile);
                LoadSettings();

                MessageBox.Show("Hei og velkommen til KGSA programmet!\n\nMen vent, vi må første se på noen innstillinger før ranking-programmet kan taes i bruk for din avdeling.\nMange standard innstillinger vil virke greit, men fyll minst ut butikk nummeret fra Elguide - denne kreves for å importere data fra Elguide.\n\nAll bruk av programmet er også på eget ansvar.\nLykke til!",
                    "KGSA - Installasjon", MessageBoxButtons.OK, MessageBoxIcon.None);
                OpenSettings();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Unntak oppstod under installasjons rutinen. Hvis problemet vedvarer reinstaller programmet.", Color.Red);
            }
        }

        private void OpenSettings()
        {
            try
            {
                if (IsBusy())
                    return;

                Logg.Status("Åpner innstillinger..");
                this.Update();
                if (lagretAvdeling > 1000)
                    appConfig.Avdeling = lagretAvdeling;
                appConfig.savedTab = tabControlMain.SelectedTab.Text;
                SaveSettings();
                var settingsForm = new FormSettings(this);
                settingsForm.StartPosition = FormStartPosition.CenterScreen;
                Logg.Status("");

                DialogResult result = settingsForm.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK || settingsForm.forceUpdate)
                {
                    processing.SetVisible = true;
                    Logg.Status("Oppdaterer..");
                    processing.SetValue = 5;
                    LoadSettings();
                    lagretAvdeling = appConfig.Avdeling;
                    database.ClearCacheTables();
                    UpdateFavorites(); // Oppdater favoritter tidligere pga mulig endring av favoritter under innstillinger. Er også under UpdateUI()
                    RetrieveDb();
                    RetrieveDbService();
                    RetrieveDbStore();
                    Logg.Status("Laster fra databasen..");
                    processing.SetValue = 50;
                    if (EmptyDatabase())
                        Reload(true); // tving oppdatering hvis databasen var tømt
                    else
                        ClearHash();
                    if (EmptyStoreDatabase())
                        ReloadStore(true); // tving oppdatering hvis lager databasen var tømt
                    else
                        ClearHashStore();
                    if (service.dbServiceDatoFra == service.dbServiceDatoTil) // tving oppdatering hvis databasen var tømt
                        ReloadService(true);
                    else
                    {
                        ClearHash("ServiceOversikt");
                        ClearHash("ServiceList");
                    }
                    Logg.Status("Fullfører endringer..");
                    processing.SetValue = 90;
                    UpdateUi();
                    processing.SetValue = 99;
                    processing.SetVisible = false;
                    Logg.Status("");
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void StartServices()
        {
            if (blueServer != null)
                blueServer.StopServer();

            if (appConfig.blueServerIsEnabled)
            {
                blueServer = new BluetoothServer(this);
                blueServer.StartServer(true);
            }
        }

        private void Reload(bool forced = false)
        {
            try
            {
                if (forced)
                {
                    ClearHash();
                    selgerkodeList.Clear();
                    comboDBselgerkode.Items.Clear();
                    listBoxSk.Items.Clear();
                    listBox_GraphSelgere.Items.Clear();
                    _graphInitialized = false;
                }
                if (EmptyDatabase())
                {
                    Logg.Log("Databasen er tom. Importer transaksjoner fra Elguide!");

                    labelRankingLastDateBig.ForeColor = SystemColors.ControlText;
                    labelRankingLastDateBig.Text = "(tom)";
                    labelRankingLastDate.ForeColor = SystemColors.ControlText;
                    labelRankingLastDate.Text = "";

                    labelGraphLastDateBig.ForeColor = SystemColors.ControlText;
                    labelGraphLastDateBig.Text = "(tom)";
                    labelGraphLastDate.ForeColor = SystemColors.ControlText;
                    labelGraphLastDate.Text = "";

                    webHTML.Navigate(htmlImport);

                    ShowHideGui_EmptyRanking(false);

                    buttonOppdater.BackColor = SystemColors.ControlLight;
                    buttonOppdater.ForeColor = SystemColors.ControlText;
                    DataTable DT = (DataTable)dataGridTransaksjoner.DataSource;
                    if (DT != null)
                        DT.Clear();
                    selgerkodeList.Clear();
                }
                else
                {
                    labelRankingLastDateBig.Text = appConfig.dbTo.ToString("dddd", norway);
                    labelRankingLastDate.Text = appConfig.dbTo.ToString("d. MMM", norway);

                    labelGraphLastDateBig.Text = appConfig.dbTo.ToString("dddd", norway);
                    labelGraphLastDate.Text = appConfig.dbTo.ToString("d. MMM", norway);

                    if ((DateTime.Now - appConfig.dbTo).Days >= 3)
                    {
                        labelRankingLastDateBig.ForeColor = Color.Red;
                        labelRankingLastDate.ForeColor = Color.Red;

                        labelGraphLastDateBig.ForeColor = Color.Red;
                        labelGraphLastDate.ForeColor = Color.Red;
                    }
                    else
                    {
                        labelRankingLastDateBig.ForeColor = SystemColors.ControlText;
                        labelRankingLastDate.ForeColor = SystemColors.ControlText;

                        labelGraphLastDateBig.ForeColor = SystemColors.ControlText;
                        labelGraphLastDate.ForeColor = SystemColors.ControlText;
                    }

                    ShowHideGui_EmptyRanking(true);

                    if (!autoMode)
                        UpdateRank();
                    highlightDate = appConfig.dbTo;
                    moveDate(0, true);
                    InitDB();
                    topgraph = new TopGraph(this);
                    sKoder = new Selgerkoder(this, true);
                    transInitialized = false;
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Kritisk feil ved initialisering av databasen.\nInstaller programmet på nytt hvis problemet vedvarer.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void ReloadBudget(bool forced = false)
        {
            try
            {
                if (forced)
                {
                    ClearBudgetHash(BudgetCategory.None);
                }
                if (EmptyDatabase() || !appConfig.experimental)
                {
                    webBudget.Navigate(htmlImport);
                    groupBudgetPages.Enabled = false;
                    groupBudgetChoices.Enabled = false;

                    buttonBudgetUpdate.BackColor = SystemColors.ControlLight;
                    buttonBudgetUpdate.ForeColor = SystemColors.ControlText;
                }
                else
                {
                    groupBudgetPages.Enabled = true;
                    groupBudgetChoices.Enabled = true;

                    budget = new BudgetObj(this);

                    if (!autoMode)
                        UpdateBudget();

                    moveBudgetDate(0, true);
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Kritisk feil ved initialisering av databasen.\nInstaller programmet på nytt hvis problemet vedvarer.", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void RunTopGraphUpdate(string katArg = "Butikk")
        {
            if (!bwUpdateTopGraph.IsBusy)
                bwUpdateTopGraph.RunWorkerAsync(katArg);
        }

        private void bwUpdateTopGraph_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string katArg = (string)e.Argument;
            if (topgraph != null)
                topgraph.UpdateGraph(katArg);
        }

        private void bwUpdateTopGraph_Completed(object sender, AsyncCompletedEventArgs e)
        {
            ProgressStop();
            graphPanelTop.Invalidate();
        }

        public bool HarSisteVersjon(string katArg, string oldHash)
        {
            try
            {
                if (datoPeriodeVelger)
                    return false;

                string html = "";
                if (katArg == "Butikk")
                    html = htmlRankingButikk;
                else if (katArg == "KnowHow")
                    html = htmlRankingKnowHow;
                else if (katArg == "Data")
                    html = htmlRankingData;
                else if (katArg == "AudioVideo")
                    html = htmlRankingAudioVideo;
                else if (katArg == "Tele")
                    html = htmlRankingTele;
                else if (katArg == "Oversikt")
                    html = htmlRankingOversikt;
                else if (katArg == "Toppselgere")
                    html = htmlRankingToppselgere;
                else if (katArg == "Lister")
                    html = htmlRankingLister;
                else if (katArg == "Vinnprodukter")
                    html = htmlRankingVinn;
                else if (katArg == "Tjenester")
                    html = htmlAvdTjenester;
                else if (katArg == "Snittpriser")
                    html = htmlAvdSnittpriser;
                else
                    return false;

                if (File.Exists(html))
                {
                    string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                    if (newHash == oldHash)
                    {
                        buttonOppdater.BackColor = SystemColors.ControlLight;
                        buttonOppdater.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: bruker mellomlagret kopi!");
                        return true;
                    }
                    else
                    {
                        buttonOppdater.BackColor = SystemColors.ControlLight;
                        buttonOppdater.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: Genererer ny ranking..");
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

        public bool HarSisteVersjonBudget(BudgetCategory cat, string oldHash)
        {
            try
            {
                if (datoPeriodeVelger)
                    return false;

                string html = "";
                if (cat == BudgetCategory.MDA)
                    html = htmlBudgetMda;
                else if (cat == BudgetCategory.AudioVideo)
                    html = htmlBudgetAudioVideo;
                else if (cat == BudgetCategory.SDA)
                    html = htmlBudgetSda;
                else if (cat == BudgetCategory.Tele)
                    html = htmlBudgetTele;
                else if (cat == BudgetCategory.Data)
                    html = htmlBudgetData;
                else if (cat == BudgetCategory.Cross)
                    html = htmlBudgetCross;
                else if (cat == BudgetCategory.Kasse)
                    html = htmlBudgetKasse;
                else if (cat == BudgetCategory.Aftersales)
                    html = htmlBudgetAftersales;
                else if (cat == BudgetCategory.Daglig)
                    html = htmlBudgetDaily;
                else
                    return false;

                if (File.Exists(html))
                {
                    string newHash = appConfig.Avdeling + pickerBudget.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                    if (newHash == oldHash)
                    {

                        buttonBudgetUpdate.BackColor = SystemColors.ControlLight;
                        buttonBudgetUpdate.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + BudgetCategoryClass.TypeToName(cat) + "]: bruker mellomlagret kopi!");
                        return true;
                    }
                    else
                    {
                        buttonBudgetUpdate.BackColor = SystemColors.ControlLight;
                        buttonBudgetUpdate.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + BudgetCategoryClass.TypeToName(cat) + "]: Genererer ny ranking..");
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

        public bool HarSisteVersjonService(string katArg, string oldHash)
        {
            try
            {
                string html = "";
                if (katArg == "ServiceOversikt")
                    html = htmlServiceOversikt;
                else if (katArg == "ServiceList")
                    html = htmlServiceList;
                else
                    return false;

                if (File.Exists(html))
                {
                    string newHash = appConfig.Avdeling + pickerRankingDate.Value.ToString() + RetrieveLinkerTimestamp().ToShortDateString();
                    if (newHash == oldHash)
                    {
                        buttonServiceOppdater.BackColor = SystemColors.ControlLight;
                        buttonServiceOppdater.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: bruker mellomlagret kopi!");
                        return true;
                    }
                    else
                    {
                        buttonServiceOppdater.BackColor = SystemColors.ControlLight;
                        buttonServiceOppdater.ForeColor = SystemColors.ControlText;
                        Logg.Debug("[" + katArg + "]: Genererer ny ranking..");
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

        /// <summary>
        /// Sjekk av databasens siste og første dato.
        /// </summary>
        /// <param name="forced">Sjekk fra til dato i databasen for transaksjoner</param>
        public void RetrieveDb(bool forced = false)
        {
            try
            {
                string dateFormat = "dd/MM/yyyy HH:mm:ss";
                appConfig.dbFrom = DateTime.Now;
                //dbTilDT = rangeMin;

                if (appConfig.dbFrom == null || appConfig.dbTo == null
                    || appConfig.dbFrom == DateTime.MinValue || appConfig.dbTo == DateTime.MinValue
                    || forced || newInstall
                    || appConfig.dbFrom > appConfig.dbTo)
                {
                    Logg.Log("Oppdaterer databasen..", null, true);

                    SqlCeCommand cmd = new SqlCeCommand("SELECT MIN(Dato) AS Expr1 FROM tblSalg WHERE (Avdeling = '" + appConfig.Avdeling + "')", connection);
                    string temp = cmd.ExecuteScalar().ToString();
                    if (!String.IsNullOrEmpty(temp))
                        appConfig.dbFrom = DateTime.ParseExact(temp, dateFormat, FormMain.norway);

                    cmd = new SqlCeCommand("SELECT MAX(Dato) AS Expr1 FROM tblSalg WHERE (Avdeling = '" + appConfig.Avdeling + "')", connection);
                    temp = cmd.ExecuteScalar().ToString();
                    if (!String.IsNullOrEmpty(temp))
                        appConfig.dbTo = DateTime.ParseExact(temp, dateFormat, FormMain.norway);

                    openXml.ClearDatabase();
                    database.ClearCacheTables(); // slett month cache tables

                    if (newInstall)
                    {
                        ClearHash("", true);
                        ClearBudgetHash(BudgetCategory.None, true);
                        newInstall = false;
                    }
                }

                if (!autoMode)
                {
                    // Oppdaterer avdelinger..
                    Logg.Log("Oppdaterer avdelinger..", null, true);
                    bwHentAvdelinger.RunWorkerAsync();
                }


                if (appConfig.dbFrom.Date != DateTime.Now.Date && appConfig.dbTo.Date != DateTime.Now.Date && appConfig.dbFrom.Date > rangeMin.Date && appConfig.dbTo.Date > rangeMin.Date)
                {
                    try
                    {
                        if (appConfig.dbFrom.Date != DateTime.Now.Date && appConfig.dbTo.Date != DateTime.Now.Date)
                        {
                            if (appConfig.dbTo.DayOfWeek == DayOfWeek.Sunday && appConfig.ignoreSunday && appConfig.dbFrom.Date != appConfig.dbTo.Date) { appConfig.dbTo = appConfig.dbTo.AddDays(-1); }

                            ChangeRankDateTimePicker(appConfig.dbTo, appConfig.dbFrom, appConfig.dbTo);
                            ChangeBudgetDateTimePicker(appConfig.dbTo, appConfig.dbFrom, appConfig.dbTo);
                            StatusInformation.Text = "Database:  " + appConfig.dbFrom.ToString("d. MMMM yyyy", norway) + "  -  " + appConfig.dbTo.ToString("d. MMMM yyyy", norway) + "  ";
                            Logg.Log("Databasen har transaksjoner mellom " + appConfig.dbFrom.ToString("d. MMMM yyyy", norway) + " og " + appConfig.dbTo.ToString("d. MMMM yyyy", norway) + " for din avdeling.", Color.Black, true);
                        }
                    }
                    catch
                    {
                        ChangeRankDateTimePicker(DateTime.Now, rangeMin, rangeMax);
                        ChangeBudgetDateTimePicker(DateTime.Now, rangeMin, rangeMax);
                        StatusInformation.Text = "Database: N/A  ";
                    }
                }
                else
                {
                    ChangeRankDateTimePicker(DateTime.Now, DateTime.Now, DateTime.Now);
                    ChangeBudgetDateTimePicker(DateTime.Now, DateTime.Now, DateTime.Now);
                    StatusInformation.Text = "Database: (tom)  ";
                }
                gc = new GraphClass(this);
            }
            catch
            {
                Logg.Log("Feil ved lasting av databasen, eller databasen er tom.", Color.Red);
                appConfig.dbFrom = DateTime.Now;
                appConfig.dbTo = DateTime.Now;
                ChangeRankDateTimePicker(DateTime.Now, rangeMin, rangeMax);
                ChangeBudgetDateTimePicker(DateTime.Now, rangeMin, rangeMax);
                StatusInformation.Text = "Database: N/A  ";
            }
        }

        public void RetrieveDbStore()
        {
            try
            {
                obsolete.Load(this);

                if (appConfig.dbStoreViewpoint > appConfig.dbStoreTo || appConfig.dbStoreViewpoint < appConfig.dbStoreFrom)
                    appConfig.dbStoreViewpoint = appConfig.dbStoreFrom;

                if (!EmptyStoreDatabase())
                {
                    try
                    {
                        if (appConfig.dbStoreTo.DayOfWeek == DayOfWeek.Sunday && appConfig.ignoreSunday && appConfig.dbStoreFrom.Date != appConfig.dbStoreTo.Date)
                            appConfig.dbStoreTo = appConfig.dbStoreTo.AddDays(-1);

                        ChangeStoreDateTimePicker(appConfig.dbStoreTo, appConfig.dbStoreFrom, appConfig.dbStoreTo);
                        Logg.Log("Lager databasen har statuser mellom " + appConfig.dbStoreFrom.ToString("d. MMMM yyyy", norway) + " og " + appConfig.dbStoreTo.ToString("d. MMMM yyyy", norway) + " for din avdeling.", Color.Black, true);
                    }
                    catch
                    {
                        ChangeStoreDateTimePicker(DateTime.Now, rangeMin, rangeMax);
                    }
                }
                else
                {
                    ChangeStoreDateTimePicker(DateTime.Now, DateTime.Now, DateTime.Now);
                }
            }
            catch(Exception ex)
            {
                Logg.Log("Feil ved lasting av lager databasen, eller databasen er tom.", Color.Red);
                Logg.Unhandled(ex);
                ChangeStoreDateTimePicker(DateTime.Now, rangeMin, rangeMax);
            }
        }

        private static bool transInitialized = false;

        /// <summary>
        /// Initialisering og nullstilling av transaksjonsvindu
        /// </summary>
        /// <param name="nullstill">Nullstill bare, ikke last inn selgerkoder på nytt</param>
        private void InitDB(bool nullstill = false)
        {
            try
            {
                if (!transInitialized)
                {
                    if (!nullstill && Loaded)
                        RefillDB();

                    if (appConfig.dbFrom.Date == appConfig.dbTo.Date)
                    {
                        pickerDBFra.MinDate = rangeMin;
                        pickerDBFra.MaxDate = rangeMax;
                        pickerDBFra.Value = DateTime.Now;
                        pickerDBTil.MinDate = rangeMin;
                        pickerDBTil.MaxDate = rangeMax;
                        pickerDBTil.Value = DateTime.Now;
                    }
                    else
                    {
                        pickerDBFra.MinDate = appConfig.dbFrom;
                        pickerDBFra.MaxDate = appConfig.dbTo;
                        pickerDBFra.Value = appConfig.dbFrom;
                        pickerDBTil.MinDate = appConfig.dbFrom;
                        pickerDBTil.MaxDate = appConfig.dbTo;
                        pickerDBTil.Value = appConfig.dbTo;
                        textBoxSearchTerm.Text = "";
                    }
                    transInitialized = true;
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Unntak ved henting av transaksjoner, eller databasen er tom.");
            }
        }

        public void SearchDB()
        {
            try
            {
                string sk = "";
                if (comboDBselgerkode.SelectedItem != null)
                    sk = comboDBselgerkode.SelectedItem.ToString();
                string vg = "";
                if (comboDBvaregruppe.SelectedItem != null)
                    vg = comboDBvaregruppe.SelectedItem.ToString();
                else
                    vg = comboDBvaregruppe.SelectedText;
                DateTime dtFra = pickerDBFra.Value;
                DateTime dtTil = pickerDBTil.Value;
                string skR = "";
                string vgR = "";
                if (!String.IsNullOrEmpty(sk) && sk != "ALLE")
                    skR = " (Selgerkode = '" + sk + "') ";

                if (!String.IsNullOrEmpty(vg) && vg != "ALLE")
                    vgR = " (Varegruppe = '" + vg + "') ";

                if (vg == "[Kitchen]")
                    vgR = " (Varegruppe LIKE '6%') ";

                if (vg == "[Data]")
                    vgR = " (Varegruppe = '580') ";

                if (vg == "[AudioVideo]")
                    vgR = " (Varegruppe = '280') ";

                if (vg == "[Tele]")
                    vgR = " (Varegruppe = '480') ";

                if (vg == "[MDA]")
                    vgR = " (Varegruppe LIKE '3%') ";

                if (vg == "[SDA]")
                    vgR = " (Varegruppe LIKE '1%') ";

                if (sk == "ALLE")
                    sk = "";

                if (vg == "ALLE")
                    vg = "";

                string arg = "";

                if (!String.IsNullOrEmpty(sk))
                    arg += " AND " + skR;

                if (!String.IsNullOrEmpty(vg))
                    arg += " AND " + vgR;

                string searchArg = "";
                var searchString = textBoxSearchTerm.Text;
                if (!String.IsNullOrEmpty(searchString))
                    searchArg = " AND (Varekode LIKE '%" + searchString + "%' OR Selgerkode LIKE '%" + searchString + "%' OR Bilagsnr LIKE '%" + searchString + "%') ";

                string cmd = "SELECT SalgID, Selgerkode, Bilagsnr, Varegruppe, Varekode, Dato, Antall, Btokr, Salgspris, Mva FROM tblSalg WHERE (Avdeling = '" +  appConfig.Avdeling + "') AND (Dato >= '" + dtFra.ToString("yyy-MM-dd") +
                    "') AND (Dato <= '" + dtTil.ToString("yyy-MM-dd") + "') " + arg + searchArg +
                    "ORDER BY CONVERT(datetime, Dato, 104) DESC";

                processing.SetValue = 75;

                DataTable dtResult = database.GetSqlDataTable(cmd);
                dtResult.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                dataGridTransaksjoner.DataSource = dtResult;

                textBoxDBCount.Text = dtResult.Rows.Count.ToString("#,##0");

                decimal result = 0;
                object r = dtResult.Compute("Sum(Antall)", null);
                if (!DBNull.Value.Equals(r)) { result = Convert.ToInt32(r); } else { result = 0; }
                textBoxDBSum.Text = result.ToString("#,##0");

                r = dtResult.Compute("Sum(Btokr)", null);
                if (!DBNull.Value.Equals(r)) { result = Convert.ToInt32(r); } else { result = 0; }
                textBoxDBBtokr.Text = result.ToString("#,##0");
                var rInntjen = result;

                r = dtResult.Compute("Sum(Salgspris)", null);
                if (!DBNull.Value.Equals(r)) { result = Convert.ToDecimal(r); } else { result = 0; }
                textBoxDBOmsetn.Text = result.ToString("#,##0");
                var rOmset = result;

                r = dtResult.Compute("Sum(SalgsprisExMva)", null);
                if (!DBNull.Value.Equals(r)) { result = Convert.ToDecimal(r); } else { result = 0; }
                textBoxDBomsetExMva.Text = result.ToString("#,##0");
                var rOmsetExMva = result;

                if (rOmset != 0)
                    textBoxDBmargin.Text = Math.Round(rInntjen / rOmsetExMva * 100, 2).ToString() + " %";
                else
                    textBoxDBmargin.Text = "0 %";
            }
            catch(Exception ex)
            {
                throw new ApplicationException("Fullførte ikke henting av transaksjoner.", ex);
            }
        }

        /// <summary>
        /// Søk opp transaksjoner.
        /// </summary>
        /// <param name="page">Side vi blir sendt fra (Butikk, Data osv)</param>
        /// <param name="month">Måned eller ikke</param>
        /// <param name="type">Type søk</param>
        /// <param name="data">Søke streng</param>
        private void SearchDB(string page, bool month = false, string type = "", string data = "")
        {
            try
            {
                if (EmptyDatabase())
                    return;

                pickerDBFra.MinDate = appConfig.dbFrom; pickerDBFra.MaxDate = appConfig.dbTo; pickerDBTil.MinDate = appConfig.dbFrom; pickerDBTil.MaxDate = appConfig.dbTo;

                pickerDBTil.Value = pickerRankingDate.Value;
                if (month) // Dato, måned eller dag
                {
                    if (GetFirstDayOfMonth(pickerRankingDate.Value) <= appConfig.dbFrom)
                        pickerDBFra.Value = appConfig.dbFrom;
                    else
                        pickerDBFra.Value = GetFirstDayOfMonth(pickerRankingDate.Value);
                    if (pickerRankingDate.Value.Month == appConfig.dbTo.Month && pickerRankingDate.Value.Year == appConfig.dbTo.Year)
                        pickerDBTil.Value = pickerRankingDate.Value;
                    else
                        pickerDBTil.Value = GetLastDayOfMonth(pickerRankingDate.Value);
                }
                else
                    pickerDBFra.Value = pickerRankingDate.Value;

                if (bwPopulateSk.IsBusy)
                    do
                    {
                        System.Threading.Thread.Sleep(50);
                        Application.DoEvents();
                    } while (bwPopulateSk.IsBusy);

                if (type == "s") // Selgerkoder
                    RefillDB(data);
                else
                    RefillDB();

                if (type == "v" || type == "s" || type == "t") // Varekoder
                {
                    if (page == "Data")
                        comboDBvaregruppe.SelectedIndex = 1;
                    if (page == "AudioVideo")
                        comboDBvaregruppe.SelectedIndex = 2;
                    if (page == "Tele")
                        comboDBvaregruppe.SelectedIndex = 3;
                }
                else if (type == "b") // Butikk kategorier
                {
                    if (data == "Computing")
                        comboDBvaregruppe.SelectedIndex = 1;
                    if (data == "AudioVideo")
                        comboDBvaregruppe.SelectedIndex = 2;
                    if (data == "Telecom")
                        comboDBvaregruppe.SelectedIndex = 3;
                    if (data == "MDA")
                        comboDBvaregruppe.SelectedIndex = 4;
                    if (data == "SDA")
                        comboDBvaregruppe.SelectedIndex = 5;
                    if (data == "Kitchen")
                        comboDBvaregruppe.SelectedIndex = 6;
                }
                else if (String.IsNullOrEmpty(type))
                {
                    if (page == "Data")
                        comboDBvaregruppe.SelectedIndex = 1;
                    else if (page == "AudioVideo")
                        comboDBvaregruppe.SelectedIndex = 2;
                    else if (page == "Tele")
                        comboDBvaregruppe.SelectedIndex = 3;
                    else
                        comboDBvaregruppe.SelectedIndex = 0;
                }

                if (type == "v")
                    textBoxSearchTerm.Text = data;
                else
                    textBoxSearchTerm.Text = "";
                SearchDB();
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Feil ved prep av søk.", ex);
            }
        }

        private void SearchDBold(string kat, bool filtrer, string searchArg)
        {
            try
            {
                if (kat == "selgerkode")
                {
                    if (pickerRankingDate.Value.Date != DateTime.Now.Date && !filtrer)
                    {
                        comboDBselgerkode.SelectedIndex = 0;
                        comboDBvaregruppe.SelectedIndex = 0;
                        textBoxSearchTerm.Text = searchArg;
                        SearchDB();
                        return;
                    }
                    else if (filtrer)
                    {
                        var index = comboDBselgerkode.FindStringExact(searchArg);
                        if (index > -1)
                        {
                            comboDBselgerkode.SelectedIndex = index;
                            SearchDB();
                            return;
                        }
                    }
                    else
                    {
                        Logg.Log("Ingenting å søke etter.", Color.Red);
                    }
                }
                else if (kat == "varegruppe")
                {
                    if (pickerRankingDate.Value.Date != DateTime.Now.Date && !filtrer)
                    {
                        comboDBselgerkode.SelectedIndex = 0;
                        comboDBvaregruppe.SelectedIndex = 0;
                        textBoxSearchTerm.Text = searchArg;
                        SearchDB();
                        return;
                    }
                    else if (filtrer)
                    {
                        var index = comboDBvaregruppe.FindStringExact(searchArg);
                        if (index > -1)
                        {
                            comboDBvaregruppe.SelectedIndex = index;
                            SearchDB();
                            return;
                        }
                    }
                    else
                    {
                        Logg.Log("Ingenting å søke etter.", Color.Red);
                    }
                }
                else if (kat == "search" && !String.IsNullOrEmpty(searchArg))
                {
                    if (pickerRankingDate.Value.Date != DateTime.Now.Date && !filtrer)
                    {
                        comboDBselgerkode.SelectedIndex = 0;
                        comboDBvaregruppe.SelectedIndex = 0;
                        textBoxSearchTerm.Text = searchArg;
                        SearchDB();
                        return;
                    }
                }
                textBoxSearchTerm.Text = "";
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Feil ved prep av søk (gammel).", ex);
            }
        }

        private void RefillDB(string sk = "")
        {
            TransPopulateSelgere();

            if (!String.IsNullOrEmpty(sk))
            {
                try
                {
                    if (comboDBselgerkode.Items.Count > 0)
                    {
                        for (int i = 0;  i < comboDBselgerkode.Items.Count; i++)
                        {
                            if (comboDBselgerkode.Items[i].ToString().Trim() == sk.Trim())
                                comboDBselgerkode.SelectedIndex = i;
                        }
                    }
                }
                catch
                {
                    Logg.Log("Feil: Fant ikke selgerkoden!");
                }
            }
            else
            {
                if (comboDBselgerkode.Items.Count > 0)
                    comboDBselgerkode.SelectedIndex = 0;
            }

            comboDBvaregruppe.Items.Clear();
            comboDBvaregruppe.Items.Add("ALLE");
            comboDBvaregruppe.Items.Add("[Data]");
            comboDBvaregruppe.Items.Add("[AudioVideo]");
            comboDBvaregruppe.Items.Add("[Tele]");
            comboDBvaregruppe.Items.Add("[MDA]");
            comboDBvaregruppe.Items.Add("[SDA]");
            comboDBvaregruppe.Items.Add("[Kitchen]");
            comboDBvaregruppe.Items.AddRange(staticVaregruppe);
            comboDBvaregruppe.Refresh();
            comboDBvaregruppe.SelectedIndex = 0;
        }

        private void TransPopulateSelgere()
        {
            try
            {
                if (transInitialized)
                    return;

                comboDBselgerkode.Items.Clear();

                if (selgerkodeList.Count == 0)
                {
                    UpdateSelgerkoderUI();
                }
                else
                {
                    comboDBselgerkode.Items.Add("ALLE");
                    comboDBselgerkode.Items.AddRange(selgerkodeList.ToArray());
                }


            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under henting av selgerkoder", ex);
                errorMsg.ShowDialog();
            }
        }

        private void GenerateFavorities()
        {
            try
            {
                Favoritter.Clear();
                Favoritter.Add(appConfig.Avdeling.ToString());
                List<string> s = new List<string> { };
                if (appConfig.favAvdeling.Length > 3)
                    s.AddRange(appConfig.favAvdeling.Split(','));
                s.Remove(appConfig.Avdeling.ToString());
                s.Remove("");
                Favoritter.AddRange(s);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void UpdateFavorites()
        {
            try
            {
                Favoritter.Clear();
                Favoritter.Add(appConfig.Avdeling.ToString());
                List<string> s = new List<string> { };
                if (appConfig.favAvdeling.Length > 3)
                    s.AddRange(appConfig.favAvdeling.Split(','));
                s.Remove(appConfig.Avdeling.ToString());
                s.Remove("");
                Favoritter.AddRange(s);

                favoritterToolStripMenuItem.DropDownItems.Clear(); // Blank ut favoritt meny

                var menuItem = new ToolStripMenuItem("Legg til..");
                favoritterToolStripMenuItem.DropDownItems.Add(menuItem);
                menuItem.DropDownItems.Add(appConfig.Avdeling + ": " + avdeling.Get(appConfig.Avdeling), null, this.FavorittClickAdd);

                favoritterToolStripMenuItem.DropDownItems.Add(new System.Windows.Forms.ToolStripSeparator());

                ToolStripItem subItem = new ToolStripMenuItem(lagretAvdeling + ": " + avdeling.Get(Convert.ToInt32(lagretAvdeling)) + " (din avdeling)");
                if (lagretAvdeling == appConfig.Avdeling)
                    subItem.Enabled = false;
                subItem.Click += new System.EventHandler(this.FavorittClick);
                favoritterToolStripMenuItem.DropDownItems.Add(subItem);

                for (int i = 1; i < Favoritter.Count; i++)
                {
                    var name = avdeling.Get(Favoritter[i]);
                    ToolStripItem subItemAvd;
                    if (name.Length > 4)
                        subItemAvd = new ToolStripMenuItem(Favoritter[i] + ": " + name);
                    else
                        subItemAvd = new ToolStripMenuItem(Favoritter[i]);
                    subItemAvd.Click += new System.EventHandler(this.FavorittClick);
                    favoritterToolStripMenuItem.DropDownItems.Add(subItemAvd);
                }


                if (!EmptyDatabase() && appConfig.avdelingerListAlle != null)
                {
                    if (appConfig.avdelingerListAlle.Count > 0)
                    {
                        arrayDbAvd = appConfig.avdelingerListAlle.ToArray();

                        List<string> menuitems = arrayDbAvd.ToList();
                        toolStripAvdeling.DropDownItems.Clear();

                        foreach (var menu in menuitems)
                        {
                            var name = avdeling.Get(Convert.ToInt32(menu));

                            ToolStripItem subItemAvd;
                            if (name.Length > 4)
                                subItemAvd = new ToolStripMenuItem(menu + ": " + name);
                            else
                                subItemAvd = new ToolStripMenuItem(menu);

                            subItemAvd.Click += new System.EventHandler(this.FavorittClick);
                            toolStripAvdeling.DropDownItems.Add(subItemAvd);
                        }
                    }
                }
                else
                {
                    toolStripAvdeling.DropDownItems.Clear();
                    ToolStripItem subItemTom = new ToolStripMenuItem("(tom)");
                    subItemTom.Enabled = false;
                    toolStripAvdeling.DropDownItems.Add(subItemTom);
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public string[] arrayDbAvd;

        private void bwHentAvdelinger_DoWork(object sender, DoWorkEventArgs e)
        {
            if (appConfig.avdelingerListAlle.Count > 0)
                arrayDbAvd = appConfig.avdelingerListAlle.ToArray();

            string[] list = listAvd();

            if (list != null)
                arrayDbAvd = list;
        }

        private void bwHentAvdelinger_Completed(object sender, AsyncCompletedEventArgs e)
        {
            List<string> menuitems = arrayDbAvd.ToList();
            toolStripAvdeling.DropDownItems.Clear();

            foreach (var menu in menuitems)
            {
                var name = avdeling.Get(Convert.ToInt32(menu));

                ToolStripItem subItem;
                if (name.Length > 4)
                    subItem = new ToolStripMenuItem(menu + ": " + name);
                else
                    subItem = new ToolStripMenuItem(menu);

                subItem.Click += new System.EventHandler(this.FavorittClick);
                toolStripAvdeling.DropDownItems.Add(subItem);
            }
        }

        private void velgCSV()
        {
            // Browse etter elguide irank.csv
            try
            {
                var fdlg = new OpenFileDialog();
                fdlg.Title = "Velg CVS-fil eksportert fra Elguide";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "All files (*.*)|*.*|CVS filer (*.csv)|*.csv";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                fdlg.Multiselect = true;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    csvFilesToImport.Clear();
                    csvFilesToImport.AddRange(fdlg.FileNames);
                    RunImport();
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void UpdateUi()
        {
            try
            {
                this.Text = "KGSA (" + avdeling.Get(appConfig.Avdeling) + ")";

                UpdateFavorites();
                InitSelgerkoder();

                // Oppdater Vis meny.
                btokrToolStripMenuItem.Checked = appConfig.kolInntjen;
                provisjonToolStripMenuItem.Checked = appConfig.kolProv;
                varekToolStripMenuItem.Checked = !appConfig.kolVarekoder;
                omsetningToolStripMenuItem.Checked = appConfig.kolSalgspris;
                favorittAvdelingerToolStripMenuItem.Checked = appConfig.favVis;
                grafikkToolStripMenuItem.Checked = appConfig.graphVis;
                rabattToolStripMenuItem.Checked = appConfig.kolRabatt;
                if (appConfig.rankingCompareLastyear > 0)
                    sammenligningToolStripMenuItem.Checked = true;
                else
                    sammenligningToolStripMenuItem.Checked = false;
                if (appConfig.rankingCompareLastmonth > 0)
                    sistMånedSammenligningToolStripMenuItem.Checked = true;
                else
                    sistMånedSammenligningToolStripMenuItem.Checked = false;

                // Show / Hide experimental features
                if (appConfig.experimental)
                {
                    if (!tabControlMain.TabPages.Contains(tabPageBudget))
                        tabControlMain.TabPages.Add(tabPageBudget);

                    lagreBudsjettPDFToolStripMenuItem.Visible = true;

                    toolStripSeparatorGetBudget.Visible = true;

                    androidAppToolStripMenuItem.Visible = true;
                    toolStripSeparator13.Visible = true;
                }
                else
                {
                    if (tabControlMain.TabPages.Contains(tabPageBudget))
                        tabControlMain.TabPages.Remove(tabPageBudget);

                    lagreBudsjettPDFToolStripMenuItem.Visible = false;

                    toolStripSeparatorGetBudget.Visible = false;

                    androidAppToolStripMenuItem.Visible = false;
                    toolStripSeparator13.Visible = false;
                }

                inkluderBudsjettMålIKveldstallToolStripMenuItem.Checked = appConfig.dailyBudgetIncludeInQuickRanking;
                oppdaterAutomatiskToolStripMenuItem.Checked = appConfig.dailyBudgetQuickRankingAutoUpdate;

                // Setup bluetooth server menues
                if (appConfig.blueServerIsEnabled)
                {
                    btAutoToolStripMenuItem.Checked = true;
                }
                if (blueServer == null)
                {
                    btOffToolStripMenuItem.Checked = true;
                    btOnToolStripMenuItem.Checked = false;
                }
                else
                {
                    btOnToolStripMenuItem.Checked = blueServer.IsOnline();
                    btOffToolStripMenuItem.Checked = !blueServer.IsOnline();
                }
                btAutoInvToolStripMenuItem.Checked = appConfig.blueInventoryAutoUpdate;
                btAutoDataToolStripMenuItem.Checked = appConfig.blueProductAutoUpdate;
                btAutoEanToolStripMenuItem.Checked = appConfig.blueEanAutoUpdate;

                // Oppdater Ny graf kontroller
                comboBox_GraphLengde.SelectedIndex = appConfig.graphLengthIndex;
                checkBox_GraphHitrateMTD.Checked = appConfig.graphHitrateMTD;
                checkBox_GraphZoom.Checked = appConfig.graphScreenZoom;
                    
                // Aktiver makro knapper og meny hvis relevante makro innstillinger er tilstede.
                if (!String.IsNullOrEmpty(appConfig.macroElguide) && !String.IsNullOrEmpty(appConfig.epostAvsender)
                    && !String.IsNullOrEmpty(appConfig.epostSMTPserver) && appConfig.epostSMTPport > 0
                    && !String.IsNullOrEmpty(appConfig.epostEmne) && !String.IsNullOrEmpty(appConfig.csvElguideExportFolder))
                {
                    try {
                        if (File.Exists(macroProgramService))
                            buttonServiceMacro.Enabled = true;
                        else
                            buttonServiceMacro.Enabled = false;
                    } catch (Exception ex) {
                        buttonServiceMacro.Enabled = false;
                        Logg.Unhandled(ex);
                    }
                    try {
                        if (File.Exists(macroProgram))
                            buttonRankingMakro.Enabled = true;
                        else
                            buttonRankingMakro.Enabled = false;
                    } catch (Exception ex) {
                        buttonRankingMakro.Enabled = false;
                        Logg.Unhandled(ex);
                    }
                    try {
                        if (File.Exists(macroProgramStore))
                            buttonLagerMakro.Enabled = true;
                        else
                            buttonLagerMakro.Enabled = false;
                    } catch (Exception ex) {
                        buttonLagerMakro.Enabled = false;
                        Logg.Unhandled(ex);
                    }
                }
                else
                    ShowHideGui_Macro(false);

                // Aktiver send ranking data meny hvis..
                if (!String.IsNullOrEmpty(appConfig.epostAvsender) && !String.IsNullOrEmpty(appConfig.epostSMTPserver)
                    && appConfig.epostSMTPport > 0 && !String.IsNullOrEmpty(appConfig.epostEmne) && !EmptyDatabase())
                {
                    toolMenuSendrank.Enabled = true;
                }

                // Aktiver timer..
                if (appConfig.autoRank && !EmptyDatabase())
                {
                    SetTimer();
                    Logg.Log("Automatisk sending av ranking aktivert. Neste automatiske utsending: " + timerNextRun.ToShortTimeString(), Color.Black, true);
                    UpdateTimer();
                }
                else
                    SetStatusInfo("timer", "ranking", "", DateTime.MinValue);

                // Aktiver Quick timer..
                if (appConfig.autoQuick && !EmptyDatabase())
                {
                    SetTimerQuick();
                    Logg.Log("Automatisk sending av kveldsranking aktivert. Neste automatiske utsending: "
                        + timerNextRunQuick.ToShortTimeString(), Color.Black, true);
                    UpdateTimerQuick();
                }
                else
                    SetStatusInfo("timer", "kveldstall", "", DateTime.MinValue);

                // Aktiver Service Auto Import timer..
                if (appConfig.AutoService)
                {
                    SetTimerService();
                    Logg.Log("Automatisk importering av servicer aktivert. Neste automatiske import: "
                        + timerNextRunService.ToShortTimeString(), Color.Black, true);
                    UpdateTimerService();
                }
                else
                    SetStatusInfo("timer", "service", "", DateTime.MinValue);

                // Aktiver AutoStore import timer..
                if (appConfig.AutoStore)
                {
                    SetTimerAutoStore();
                    Logg.Log("Automatisk importering av lager aktivert. Neste automatiske import: "
                        + timerNextRunAutoStore.ToShortTimeString(), Color.Black, true);
                    UpdateTimerAutoStore();
                }
                else
                    SetStatusInfo("timer", "lager", "", DateTime.MinValue);

                if (!EmptyDatabase())
                {
                    SetStatusInfo("db", "ranking", "Transaksjoner fra "
                        + appConfig.dbFrom.ToString("d. MMMM yyyy", norway) + " til "
                        + appConfig.dbTo.ToString("d. MMMM yyyy", norway), appConfig.dbTo);
                }
                else
                    SetStatusInfo("db", "ranking", "", DateTime.MinValue);
                if (!EmptyStoreDatabase())
                {
                    SetStatusInfo("db", "lager", "Lager database var sist oppdatert "
                        + appConfig.dbStoreTo.ToString("d. MMMM yyyy", norway), appConfig.dbStoreTo);
                }
                else
                    SetStatusInfo("db", "lager", "", DateTime.MinValue);
                if (service.dbServiceDatoFra != service.dbServiceDatoTil)
                {
                    SetStatusInfo("db", "service", "Service database var sist oppdatert "
                        + service.dbServiceDatoTil.ToString("d. MMMM yyyy", norway), service.dbServiceDatoTil);
                }
                else
                    SetStatusInfo("db", "service", "", DateTime.MinValue);

                if (appConfig.webserverEnabled && appConfig.webserverPort > 0
                    && appConfig.webserverPort <= 65535 && !String.IsNullOrEmpty(appConfig.webserverHost))
                {
                    if (server != null)
                    {
                        if (server.ws.IsOnline())
                            Logg.Log("Webserver er aktivert og lytter på: http://"
                                + appConfig.webserverHost + ":" + appConfig.webserverPort + "/", Color.Green);
                        else
                            MessageBox.Show("Obs! Webserver er aktivert men lytter ikke til angitt adresse ("
                                + appConfig.webserverHost + ")\nSjekk innstillinger og/eller brannmur."
                                + " Prøv evt. å starte programmet som administrator.",
                                "KGSA - Advarsel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                if (appConfig.histogramVis)
                {
                    panel4.Visible = true;
                    toolStripMenuItemVisHist.Checked = true;
                    string kat = currentPage();
                    if (kat == "Butikk" || kat == "Data" || kat == "AudioVideo" || kat == "Tele")
                        RunTopGraphUpdate(kat);
                    else
                        RunTopGraphUpdate();
                }
                else
                {
                    panel4.Visible = false;
                    toolStripMenuItemVisHist.Checked = false;
                }

                if (appConfig.importSetting.StartsWith("Full"))
                    ShowHideGui_FullTrans(true);
                else
                    ShowHideGui_FullTrans(false);

                if (!String.IsNullOrEmpty(appConfig.savedTab))
                    foreach (TabPage tab in tabControlMain.TabPages)
                        if (tab.Text == appConfig.savedTab)
                            tabControlMain.SelectedTab = tab;
            }
            catch(Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under oppdatering av UI", ex);
                errorMsg.ShowDialog(this);
            }
        }
        private void SetTimerAutoStore()
        {
            timerAutoStore.Interval = 60 * 1000; // 1 minutt
            timerAutoStore.Enabled = true;
            timerAutoStore.Start();
            timerNextRunAutoStore = DateTime.Now;
            DateTime s = timerNextRunAutoStore;

            var ts = new TimeSpan(appConfig.AutoStoreHour, (appConfig.AutoStoreMinute * 10), 0);
            timerNextRunAutoStore = s.Date + ts;

            if (timerNextRunAutoStore < DateTime.Now)
                timerNextRunAutoStore = timerNextRunAutoStore.AddDays(1);

            if (timerNextRunAutoStore.DayOfWeek == DayOfWeek.Monday && appConfig.ignoreSunday)
                timerNextRunAutoStore = timerNextRunAutoStore.AddDays(1);
        }

        private void SetTimerService()
        {
            try
            {
                timerAutoService.Interval = 60 * 1000; // 1 minutt
                timerAutoService.Enabled = true;
                timerAutoService.Start();
                timerNextRunService = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, 0, 0);
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
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void SetTimer()
        {
            timerAutoRanking.Interval = 60 * 1000; // 1 minutt
            timerAutoRanking.Enabled = true;
            timerAutoRanking.Start();
            timerNextRun = DateTime.Now;
            DateTime s = timerNextRun;

            var ts = new TimeSpan(appConfig.epostHour, (appConfig.epostMinute * 10), 0);
            timerNextRun = s.Date + ts;

            if (timerNextRun < DateTime.Now)
                timerNextRun = timerNextRun.AddDays(1);

            if (timerNextRun.DayOfWeek == DayOfWeek.Monday && appConfig.ignoreSunday)
                timerNextRun = timerNextRun.AddDays(1);
        }

        private void SetTimerQuick()
        {
            timerAutoQuick.Interval = 60 * 1000; // 1 minutt
            timerAutoQuick.Enabled = true;
            timerAutoQuick.Start();
            timerNextRunQuick = DateTime.Now;
            DateTime s = timerNextRunQuick;

            var ts = new TimeSpan(appConfig.epostHourQuick, (appConfig.epostMinuteQuick * 10), 0);
            if (timerNextRunQuick < DateTime.Now)
            {
                if (s.AddDays(1).DayOfWeek == DayOfWeek.Saturday && appConfig.epostTimerQuickSaturday)
                    ts = new TimeSpan(appConfig.epostHourQuickSaturday, (appConfig.epostMinuteQuickSaturday * 10), 0);
                if (s.AddDays(1).DayOfWeek == DayOfWeek.Sunday && appConfig.epostTimerQuickSunday && !appConfig.ignoreSunday)
                    ts = new TimeSpan(appConfig.epostHourQuickSunday, (appConfig.epostMinuteQuickSunday * 10), 0);
            }
            else
            {
                if (s.DayOfWeek == DayOfWeek.Saturday && appConfig.epostTimerQuickSaturday)
                    ts = new TimeSpan(appConfig.epostHourQuickSaturday, (appConfig.epostMinuteQuickSaturday * 10), 0);
                if (s.DayOfWeek == DayOfWeek.Sunday && appConfig.epostTimerQuickSunday && !appConfig.ignoreSunday)
                    ts = new TimeSpan(appConfig.epostHourQuickSunday, (appConfig.epostMinuteQuickSunday * 10), 0);
            }
            timerNextRunQuick = s.Date + ts;

            if (timerNextRunQuick < DateTime.Now)
                timerNextRunQuick = timerNextRunQuick.AddDays(1);

            if (timerNextRunQuick.DayOfWeek == DayOfWeek.Sunday && appConfig.ignoreSunday)
                timerNextRunQuick = timerNextRunQuick.AddDays(1);
        }

        public void UpdateTimerAutoStore()
        {
            try
            {
                if (appConfig.AutoStore)
                {
                    TimeSpan check = timerNextRunAutoStore.Subtract(DateTime.Now);
                    SetStatusInfo("timer", "lager", "", timerNextRunAutoStore);

                    if (check.TotalMinutes < 1 && check.TotalMinutes >= 0)
                    {
                        if (!IsBusy(true))
                            delayedAutoStore();
                        return;
                    }

                    if (check.TotalMinutes < 0)
                    {
                        DateTime s = timerNextRunAutoStore;
                        var ts = new TimeSpan(appConfig.AutoStoreHour, (appConfig.AutoStoreMinute * 10), 0);
                        timerNextRunAutoStore = s.Date.AddDays(1) + ts;
                        if (timerNextRunAutoStore.DayOfWeek == DayOfWeek.Sunday)
                            timerNextRunAutoStore = timerNextRunAutoStore.AddDays(1);
                        return;
                    }

                    if (check.TotalMinutes < 15 && check.TotalMinutes > 1)
                    {
                        if (Math.Round(check.TotalMinutes) == 1)
                            Logg.Log("Starter automatisk innhenting av lager om 1 minutt.");
                        else
                            Logg.Log("Starter automatisk innhenting av lager om " + Math.Round(check.TotalMinutes - 1) + " minutter.");
                        return;
                    }
                }
                else
                {
                    timerAutoStore.Stop();
                    Logg.Log("Automatisk innhenting av transaksjoner deaktivert.");
                }
            }
            catch
            {
                Logg.Log("Unntak i timer-funksjon. Omstart av programmet anbefales.", Color.Red);
            }
        }

        public void UpdateTimer()
        {
            try
            {
                if (appConfig.autoRank)
                {
                    TimeSpan check = timerNextRun.Subtract(DateTime.Now);
                    SetStatusInfo("timer", "ranking", "", timerNextRun);

                    if (check.TotalMinutes < 1 && check.TotalMinutes >= 0)
                    {
                        if (!IsBusy(true))
                        {
                            RestoreWindow();
                            bwAutoRanking.RunWorkerAsync(); // Starter jobb som starter makro, importering, ranking, pdf konvertering og sending på mail.
                            processing.SetVisible = true;
                            while (bwAutoRanking.IsBusy)
                            {
                                Application.DoEvents();
                                System.Threading.Thread.Sleep(100);
                            }
                            processing.HideDelayed();
                            this.Activate();
                        }
                        return;
                    }

                    if (check.TotalMinutes < 0)
                    {
                        DateTime s = timerNextRun;
                        var ts = new TimeSpan(appConfig.epostHour, (appConfig.epostMinute * 10), 0);
                        timerNextRun = s.Date.AddDays(1) + ts;
                        if (timerNextRun.DayOfWeek == DayOfWeek.Sunday)
                            timerNextRun = timerNextRun.AddDays(1);
                        return;
                    }

                    if (check.TotalMinutes < 15 && check.TotalMinutes > 1)
                    {
                        if (Math.Round(check.TotalMinutes) == 1)
                            Logg.Log("Starter automatisk innhenting av transaksjoner om 1 minutt.");
                        else
                            Logg.Log("Starter automatisk innhenting av transaksjoner om " + Math.Round(check.TotalMinutes) + " minutter.");
                        return;
                    }
                }
                else
                {
                    timerAutoRanking.Stop();
                    Logg.Log("Automatisk innhenting av transaksjoner deaktivert.");
                }
            }
            catch
            {
                Logg.Log("Unntak i timer-funksjon. Omstart av programmet anbefales.", Color.Red);
            }
        }

        public void UpdateTimerQuick()
        {
            try
            {
                if (appConfig.autoQuick)
                {
                    TimeSpan check = timerNextRunQuick.Subtract(DateTime.Now);
                    SetStatusInfo("timer", "kveldstall", "", timerNextRunQuick);

                    if (check.TotalMinutes < 1 && check.TotalMinutes >= 0)
                    {
                        if (!IsBusy(true))
                        {
                            RestoreWindow();
                            bwQuickAuto.RunWorkerAsync(); // Starter jobb som starter makro, importering, ranking, pdf konvertering og sending på mail.
                            processing.SetVisible = true;
                            while (bwQuickAuto.IsBusy)
                            {
                                Application.DoEvents();
                                System.Threading.Thread.Sleep(100);
                            }
                            processing.HideDelayed();
                            this.Activate();
                        }
                        return;
                    }

                    if (check.TotalMinutes < 0)
                    {
                        DateTime s = timerNextRunQuick;
                        var ts = new TimeSpan(appConfig.epostHourQuick, (appConfig.epostMinuteQuick * 10), 0);
                        timerNextRunQuick = s.Date.AddDays(1) + ts;
                        if (timerNextRunQuick.DayOfWeek == DayOfWeek.Sunday)
                            timerNextRunQuick = timerNextRunQuick.AddDays(1);
                        return;
                    }

                    if (check.TotalMinutes < 15 && check.TotalMinutes > 1)
                    {
                        if (Math.Round(check.TotalMinutes) == 1)
                            Logg.Log("Starter automatisk innhenting av kveldstall om 1 minutt.");
                        else
                            Logg.Log("Starter automatisk innhenting av kveldstall om " + Math.Round(check.TotalMinutes - 1) + " minutter.");
                        return;
                    }
                }
                else
                {
                    timerAutoQuick.Stop();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void exportDatabase()
        {
            try
            {
                if (IsBusy(true))
                    return;

                SaveFileDialog saveFileDB = new SaveFileDialog();
                saveFileDB.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                saveFileDB.Filter = "Compact database (*.SDF)|*.sdf|Alle filer (*.*)|*.*";
                saveFileDB.FilterIndex = 1;
                saveFileDB.FileName = "KGSA " + databaseVersion + " " + appConfig.dbFrom.ToString("yyMMd") + "-" + appConfig.dbTo.ToString("yyMMd");

                if (saveFileDB.ShowDialog() == DialogResult.OK)
                {
                    processing.SetVisible = true;
                    processing.SetText = "Eksporterer databasen..";
                    
                    try
                    {
                        if (File.Exists(saveFileDB.FileName))
                        {
                            Logg.Log("File eksisterer allerede, skriver over..", Color.Black, true);
                            File.Delete(saveFileDB.FileName);
                        }
                        processing.SetText = "Eksporterer databasen..";
                        File.Copy(settingsPath + @"\" + databaseName, saveFileDB.FileName);
                        Logg.Log("Fullført eksportering av databasen. (" + saveFileDB.FileName + ")", Color.Green);
                        processing.SetText = "Ferdig!";
                    }
                    catch (IOException ex)
                    {
                        FormError errorMsg = new FormError("Filen er i bruk eller ble nektet tilgang.", ex);
                        errorMsg.ShowDialog(this);
                    }
                    catch (Exception ex)
                    {
                        FormError errorMsg = new FormError("Ukjent feil oppstod under eksportering av databasen.\nOperasjon avbrutt.", ex);
                        errorMsg.ShowDialog(this);
                    }
                }
            }
            catch
            {
                Logg.Log("Feil under eksportering av databasen.", Color.Red);
            }
        }

        private void importDatabase()
        {
            try
            {
                if (IsBusy())
                    return;

                DialogResult iAmSure = MessageBox.Show("Importering av en database bytter ut den eksisterende!\nAlle transaksjoner, lagerstatus, e-post adresser og selgerkoder i den gjeldende databasen blir slettet\n\n\nEr du sikker på du vil fortsette?", "KGSA - Advarsel" , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information,MessageBoxDefaultButton.Button2);
                if (iAmSure == DialogResult.Yes)
                {
                    OpenFileDialog importFileDB = new OpenFileDialog();
                    importFileDB.InitialDirectory = Convert.ToString(Environment.SpecialFolder.MyDocuments);
                    importFileDB.Filter = "Compact database (*.SDF)|*.sdf";
                    importFileDB.FilterIndex = 1;

                    if (importFileDB.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            if (!importFileDB.FileName.ToLower().Contains(databaseVersion.ToLower()))
                                if (MessageBox.Show("Databasen kan være inkompatibel med denne versjonen av KGSA!\nGjeldene versjon: " + databaseVersion + "\n\nSikker på at du vil fortsette?", "KGSA - Advarsel", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                                    return;

                            processing.SetVisible = true;
                            processing.SetText = "Importerer databasen..";

                            string sqlSource = @"Data Source=" + importFileDB.FileName;

                            if (!connection.TableExists("tblSelgerkoder") || !connection.FieldExists("tblSalg", "Mva"))
                            {
                                Logg.Log("Databasen (" + importFileDB.FileName + ") er ikke kompatibel med denne versjonen av KGSA.", Color.Red);
                                MessageBox.Show("Databasen er ikke kompatibel med denne versjonen av KGSA.\nImportering avbrutt!", "KGSA - Database ikke kompatibel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                processing.SetText = "Importering avbrutt..";
                                return;
                            }

                            Logg.Log("Importerer ny database (" + importFileDB.FileName + ") ..");
                            try
                            {
                                database.CloseConnection();
                                processing.SetValue = 20;
                                if (File.Exists(settingsPath + @"\DatabaseCopy.sdf"))
                                    File.Delete(settingsPath + @"\DatabaseCopy.sdf");
                                File.Copy(importFileDB.FileName, settingsPath + @"\DatabaseCopy.sdf");
                                processing.SetValue = 60;
                                if (File.Exists(settingsPath + @"\DatabaseBackup.sdf"))
                                    File.Delete(settingsPath + @"\DatabaseBackup.sdf");

                                File.Replace(settingsPath + @"\DatabaseCopy.sdf", fileDatabase, settingsPath + @"\DatabaseBackup.sdf");
                                processing.SetText = "Kopiering fullført. Initialiserer..";
                                processing.SetValue = 70;
                                database.OpenConnection();
                                database.VerifyDatabase();
                                RetrieveDbStore();
                                RetrieveDb(true);
                                processing.SetValue = 80;
                                ReloadStore(true);
                                Reload(true);
                                ReloadBudget(true);
                                RunTopGraphUpdate();
                                processing.SetValue = 95;
                                UpdateUi();
                                if (File.Exists(settingsPath + @"\DatabaseBackup.sdf"))
                                    File.Delete(settingsPath + @"\DatabaseBackup.sdf");
                                processing.SetText = "Ferdig!";
                            }
                            catch (IOException ex)
                            {
                                FormError errorMsg = new FormError("Ingen tilgang", ex, "Filen/databasen er i bruk eller ingen tilgang.");
                                errorMsg.ShowDialog(this);
                            }
                            catch(Exception ex)
                            {
                                Logg.Unhandled(ex);
                                Logg.Log("Ukjent feil oppstod under importering av databasen. Se logg for detaljer.", Color.Red);
                            }
                            finally
                            {
                                if (File.Exists(settingsPath + @"\DatabaseCopy.sdf"))
                                    File.Delete(settingsPath + @"\DatabaseCopy.sdf");
                            }
                        }
                        catch (SqlCeException ex)
                        {
                            FormError errorMsg = new FormError("Feil under lesing av databasen.", ex, "Databasen er i feil versjon, er skadet eller programmet ble nektet tilgang.");
                            errorMsg.ShowDialog(this);
                        }
                        catch (Exception ex)
                        {
                            FormError errorMsg = new FormError("Ukjent feil oppstod.", ex, "Feil under innhenting av ny database.");
                            errorMsg.ShowDialog(this);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }
    }

    public static class SqlCeExtentions
    {
        public static bool TableExists(this SqlCeConnection connection, string tableName)
        {
            if (tableName == null) throw new ArgumentNullException("tableName");
            if (tableName.IsNullOrWhiteSpace()) throw new ArgumentException("Invalid table name");
            if (connection == null) throw new ArgumentNullException("connection");
            if (connection.State != ConnectionState.Open)
            {
                throw new InvalidOperationException("TableExists requires an open and available Connection. The connection's current state is " + connection.State);
            }

            using (SqlCeCommand command = connection.CreateCommand())
            {
                command.CommandType = CommandType.Text;
                command.CommandText = "SELECT 1 FROM Information_Schema.Tables WHERE TABLE_NAME = @tableName";
                command.Parameters.AddWithValue("tableName", tableName);
                object result = command.ExecuteScalar();
                return result != null;
            }
        }

        public static bool FieldExists(this SqlCeConnection connection, string tableName, string fieldName)
        {
            if (tableName == null) throw new ArgumentNullException("tableName");
            if (fieldName == null) throw new ArgumentNullException("fieldName");
            if (tableName.IsNullOrWhiteSpace()) throw new ArgumentException("Invalid table name");
            if (connection == null) throw new ArgumentNullException("connection");
            if (connection.State != ConnectionState.Open)
            {
                throw new InvalidOperationException("TableExists requires an open and available Connection. The connection's current state is " + connection.State);
            }

            var tblQuery = "SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS"
                + " WHERE TABLE_NAME = @tableName AND"
                + " COLUMN_NAME = @columnName";

            using (SqlCeCommand command = connection.CreateCommand())
            {
                command.CommandText = tblQuery;
                var tblNameParam = new SqlCeParameter(
                    "@tableName",
                    SqlDbType.NVarChar,
                    128);

                tblNameParam.Value = tableName;
                command.Parameters.Add(tblNameParam);
                var colNameParam = new SqlCeParameter(
                    "@columnName",
                    SqlDbType.NVarChar,
                    128);

                colNameParam.Value = fieldName;
                command.Parameters.Add(colNameParam);
                object objvalid = command.ExecuteScalar(); // will return 1 or null
                return objvalid != null;
            }
        }
    }
}