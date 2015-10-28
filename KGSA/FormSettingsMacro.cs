using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormSettingsMacro : Form
    {
        private readonly Timer timerMsgClear = new Timer();
        FormMain main;
        KgsaEmail kgsaEmail;
        bool statusElguide = false;
        bool statusEpost = false;
        public FormSettingsMacro(FormMain form)
        {
            this.main = form;
            this.kgsaEmail = new KgsaEmail(main);
            InitializeComponent();
            timerMsgClear.Tick += timer;
            Log.d("Makro Innstillinger åpnet.");
        }

        private void timer(object sender, EventArgs e)
        {
            ClearMessageTimerStop();
        }

        private void ClearMessageTimer()
        {
            timerMsgClear.Stop();
            timerMsgClear.Interval = 10 * 1000; // 10 sek
            timerMsgClear.Enabled = true;
            timerMsgClear.Start();
        }

        private void ClearMessageTimerStop()
        {
            if (timerMsgClear.Enabled)
            {
                SendMessage("");
                timerMsgClear.Stop();
                timerMsgClear.Enabled = false;
            }
        }

        private void SetStatusInfo(string kat, string type, string str, DateTime date)
        {
            try
            {
                if (kat == "timer")
                {
                    if (date == DateTime.MinValue)
                    {
                        str = "Avslått";
                    }
                    else
                    {
                        TimeSpan ts = date.Subtract(DateTime.Now);
                        str = str + date.ToShortDateString() + " " + date.ToShortTimeString() + " (om " + ToReadableString(ts) + ")";
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
            catch (Exception ex)
            {
                SendMessage(ex.Message.ToString().Replace("\n", "").Replace("\r", ""), Color.Red);
                Log.Unhandled(ex);
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

        private bool ImportMacroSettings()
        {
            try
            {

                // Fane: MAKRO
                // MAKRO - Generelt
                textBoxElguide.Text = main.appConfig.macroElguide;
                numericAutoDelay.Value = main.appConfig.macroLatency;
                checkBoxMacroWarning.Checked = main.appConfig.macroShowWarning;

                // MAKRO - Makro program - Automatisk ranking
                checkBoxEpostTimer.Checked = main.appConfig.autoRank;
                comboBoxEpostTime.SelectedIndex = main.appConfig.epostHour;
                comboBoxEpostMinutt.SelectedIndex = main.appConfig.epostMinute;
                checkEpostOnlySendUpdated.Checked = main.appConfig.epostOnlySendUpdated;
                textBoxEpostEmne.Text = main.appConfig.epostEmne;
                textBoxEpostBody.Text = main.appConfig.epostBody.Replace("\n", Environment.NewLine);

                // MAKRO - Makro program - Automatisk service
                checkBoxAutoService.Checked = main.appConfig.AutoService;
                SetNumericValue(numericServiceRecurringMinute, main.appConfig.serviceAutoImportMinutter);
                comboBoxServiceTimeFrom.SelectedIndex = main.appConfig.serviceAutoImportFraIndex;
                comboBoxServiceTimeTo.SelectedIndex = main.appConfig.serviceAutoImportTilIndex;

                // MAKRO - Makro program - Automatisk kveldstall
                checkBoxSendQuickRanking.Checked = main.appConfig.autoQuick;
                comboBoxQuickHour.SelectedIndex = main.appConfig.epostHourQuick;
                comboBoxQuickMinute.SelectedIndex = main.appConfig.epostMinuteQuick;
                textBoxQuickTitle.Text = main.appConfig.epostEmneQuick;
                textBoxQuickBody.Text = main.appConfig.epostBodyQuick.Replace("\n", Environment.NewLine);
                checkBoxAutoServiceIncRank.Checked = main.appConfig.epostIncService;

                checkBoxQuickOtherTimeSaturday.Checked = main.appConfig.epostTimerQuickSaturday;
                checkBoxQuickOtherTimeSunday.Checked = main.appConfig.epostTimerQuickSunday;
                comboBoxQuickHourSaturday.SelectedIndex = main.appConfig.epostHourQuickSaturday;
                comboBoxQuickMinuteSaturday.SelectedIndex = main.appConfig.epostMinuteQuickSaturday;
                comboBoxQuickHourSunday.SelectedIndex = main.appConfig.epostHourQuickSunday;
                comboBoxQuickMinuteSunday.SelectedIndex = main.appConfig.epostMinuteQuickSunday;
                checkMacroImportQuickSales.Checked = main.appConfig.macroImportQuickSales;

                // MAKRO - Lager makro program
                checkBoxAutoStore.Checked = main.appConfig.AutoStore;
                comboBoxAutoStoreHour.SelectedIndex = main.appConfig.AutoStoreHour;
                comboBoxAutoStoreMinute.SelectedIndex = main.appConfig.AutoStoreMinute;
                checkBoxAutoStoreOnline.Checked = main.appConfig.onlineImporterAuto;

                // MAKRO - Makrop program - Program (alle 4)
                try
                {
                    textBoxMacroProgram.Text = File.ReadAllText(FormMain.macroProgram);
                }
                catch
                {
                    SendMessage("Feil ved lesing av makro program.", Color.Red);
                }
                try
                {
                    textBoxMacroQuick.Text = File.ReadAllText(FormMain.macroProgramQuick);
                }
                catch
                {
                    SendMessage("Feil ved lesing av quick makro program.", Color.Red);
                }
                try
                {
                    textBoxMacroServiceProgram.Text = File.ReadAllText(FormMain.macroProgramService);
                }
                catch
                {
                    SendMessage("Feil ved lesing av Service makro program.", Color.Red);
                }
                try
                {
                    textBoxMacroStore.Text = File.ReadAllText(FormMain.macroProgramStore);
                }
                catch
                {
                    SendMessage("Feil ved lesing av lager makro program.", Color.Red);
                }

                if (main.appConfig.autoRank)
                    SetStatusInfo("timer", "ranking", "", main.timerNextRun);
                else
                    SetStatusInfo("timer", "ranking", "", DateTime.MinValue);

                if (main.appConfig.autoQuick)
                    SetStatusInfo("timer", "kveldstall", "", main.timerNextRunQuick);
                else
                    SetStatusInfo("timer", "kveldstall", "", DateTime.MinValue);

                if (main.appConfig.AutoService)
                    SetStatusInfo("timer", "service", "", main.timerNextRunService);
                else
                    SetStatusInfo("timer", "service", "", DateTime.MinValue);

                if (main.appConfig.AutoStore)
                    SetStatusInfo("timer", "lager", "", main.timerNextRunAutoStore);
                else
                    SetStatusInfo("timer", "lager", "", DateTime.MinValue);

                try
                {
                    if (!String.IsNullOrEmpty(main.appConfig.epostAvsender) && !String.IsNullOrEmpty(main.appConfig.epostSMTPserver) &&
                        main.appConfig.epostSMTPport > 0)
                    {
                        if (kgsaEmail.emailDb.Rows.Count == 0)
                        {
                            labelEpostStatus.Text = "Ingen e-post adresser lagret.";
                            labelEpostStatus.ForeColor = Color.Red;
                            this.statusEpost = false;
                        }
                        else
                        {
                            labelEpostStatus.Text = "OK";
                            labelEpostStatus.ForeColor = Color.Green;
                            this.statusEpost = true;
                        }
                    }
                    else
                    {
                        labelEpostStatus.Text = "Mangler innstillinger.";
                        labelEpostStatus.ForeColor = Color.Red;
                        this.statusEpost = false;
                    }
                }
                catch
                {
                    labelEpostStatus.Text = "Sjekk e-post innstillinger.";
                    labelEpostStatus.ForeColor = Color.Red;
                    this.statusEpost = false;
                }

                try
                {
                    if (File.Exists(main.appConfig.macroElguide) && Directory.Exists(main.appConfig.csvElguideExportFolder))
                    {
                        labelElguideStatus.Text = "OK";
                        labelElguideStatus.ForeColor = Color.Green;
                        this.statusElguide = true;
                    }
                    else
                    {
                        if (!File.Exists(main.appConfig.macroElguide))
                            labelElguideStatus.Text = "Mangler profil";
                        else
                            labelElguideStatus.Text = "Mappe utilg.";
                        labelElguideStatus.ForeColor = Color.Red;
                        this.statusElguide = false;
                    }
                }
                catch
                {

                }

                if (FormMain.autoMode)
                {
                    labelStatusMacro.Text = "Kjører..";
                    labelStatusMacro.ForeColor = Color.Black;
                }
                else
                {
                    labelStatusMacro.Text = "N/A";
                    labelStatusMacro.ForeColor = Color.Black;
                }

                if (statusElguide && statusEpost)
                    buttonTestMacro.Enabled = true;
                else
                    buttonTestMacro.Enabled = false;

                return true;
            }
            catch (Exception ex)
            {
                SendMessage(ex.Message.ToString().Replace("\n", "").Replace("\r", ""), Color.Red);
                Log.Unhandled(ex);
            }
            return false;
        }

        private bool ExportMacroSettings()
        {
            try
            {
                // Fane: MAKRO
                // MAKRO - Generelt
                try
                {
                    var laten = numericAutoDelay.Value;
                    if (laten < 3 && laten >= 0.5M)
                        main.appConfig.macroLatency = laten;
                    else
                    {
                        numericAutoDelay.Value = 1;
                        SendMessage("Formatfeil på makro forsink.", Color.Red);
                        throw new System.ArgumentException("Formatfeil på makro forsink.");
                    }
                }
                catch
                {
                    numericAutoDelay.Value = 1;
                    throw new System.ArgumentException("Formatfeil på makro forsink.");
                }
                main.appConfig.macroShowWarning = checkBoxMacroWarning.Checked;
                main.appConfig.macroElguide = textBoxElguide.Text;

                // MAKRO - Makro program - Automatisk ranking
                main.appConfig.autoRank = checkBoxEpostTimer.Checked;
                main.appConfig.epostHour = comboBoxEpostTime.SelectedIndex;
                main.appConfig.epostMinute = comboBoxEpostMinutt.SelectedIndex;
                main.appConfig.epostEmne = textBoxEpostEmne.Text;
                main.appConfig.epostBody = textBoxEpostBody.Text.Replace(Environment.NewLine, "\n");
                main.appConfig.epostOnlySendUpdated = checkEpostOnlySendUpdated.Checked;
                main.appConfig.epostIncService = checkBoxAutoServiceIncRank.Checked;

                // MAKRO - Makro program - Automatisk service
                main.appConfig.AutoService = checkBoxAutoService.Checked;
                main.appConfig.serviceAutoImportMinutter = (int)numericServiceRecurringMinute.Value;
                if (comboBoxServiceTimeFrom.SelectedIndex < comboBoxServiceTimeTo.SelectedIndex)
                {
                    main.appConfig.serviceAutoImportFraIndex = comboBoxServiceTimeFrom.SelectedIndex;
                    main.appConfig.serviceAutoImportTilIndex = comboBoxServiceTimeTo.SelectedIndex;
                }

                // MAKRO - Makro program - Automatisk kveldstall
                main.appConfig.autoQuick = checkBoxSendQuickRanking.Checked;
                main.appConfig.epostHourQuick = comboBoxQuickHour.SelectedIndex;
                main.appConfig.epostMinuteQuick = comboBoxQuickMinute.SelectedIndex;
                main.appConfig.epostEmneQuick = textBoxQuickTitle.Text;
                main.appConfig.epostBodyQuick = textBoxQuickBody.Text.Replace(Environment.NewLine, "\n");
                main.appConfig.epostTimerQuickSaturday = checkBoxQuickOtherTimeSaturday.Checked;
                main.appConfig.epostTimerQuickSunday = checkBoxQuickOtherTimeSunday.Checked;
                main.appConfig.epostHourQuickSaturday = comboBoxQuickHourSaturday.SelectedIndex;
                main.appConfig.epostHourQuickSunday = comboBoxQuickHourSunday.SelectedIndex;
                main.appConfig.epostMinuteQuickSaturday = comboBoxQuickMinuteSaturday.SelectedIndex;
                main.appConfig.epostMinuteQuickSunday = comboBoxQuickMinuteSunday.SelectedIndex;
                main.appConfig.macroImportQuickSales = checkMacroImportQuickSales.Checked;

                // MAKRO - Makro program Automatisk lager import
                main.appConfig.AutoStore = checkBoxAutoStore.Checked;
                main.appConfig.AutoStoreHour = comboBoxAutoStoreHour.SelectedIndex;
                main.appConfig.AutoStoreMinute = comboBoxAutoStoreMinute.SelectedIndex;
                main.appConfig.onlineImporterAuto = checkBoxAutoStoreOnline.Checked;

                // MAKRO - Makro program - Program (alle 3)
                try
                {
                    File.WriteAllText(FormMain.macroProgram, textBoxMacroProgram.Text);
                }
                catch
                {
                    SendMessage("Feil oppstod ved lagring av ranking program.");
                }
                try
                {
                    File.WriteAllText(FormMain.macroProgramQuick, textBoxMacroQuick.Text);
                }
                catch
                {
                    SendMessage("Feil oppstod ved lagring av kvelds ranking program.");
                }
                try
                {
                    File.WriteAllText(FormMain.macroProgramService, textBoxMacroServiceProgram.Text);
                }
                catch
                {
                    SendMessage("Feil oppstod ved lagring av Servoce program.");
                }
                try
                {
                    File.WriteAllText(FormMain.macroProgramStore, textBoxMacroStore.Text);
                }
                catch
                {
                    SendMessage("Feil oppstod ved lagring av lager program.");
                }


                return true;
            }
            catch (Exception ex)
            {
                SendMessage(ex.Message.ToString().Replace("\n", "").Replace("\r", ""), Color.Red);
                Log.Unhandled(ex);
            }
            return false;
        }

        private void SetNumericValue(NumericUpDown control, float value)
        {
            SetNumericValue(control, Convert.ToDecimal(value));
        }

        private void SetNumericValue(NumericUpDown control, int value)
        {
            SetNumericValue(control, Convert.ToDecimal(value));
        }

        private void SetNumericValue(NumericUpDown control, decimal value)
        {
            try
            {
                if (value >= control.Minimum && value <= control.Maximum)
                    control.Value = value;
            }
            catch { control.Value = control.Minimum; }
        }

        delegate void SetSendMessageCallback(string str, Color? c = null, bool noLog = false);

        public void SendMessage(string str, Color? c = null, bool noLog = false)
        {
            try
            {
                if (this.errorMessage.InvokeRequired)
                {
                    if (!IsHandleCreated)
                        return;

                    SetSendMessageCallback d = new SetSendMessageCallback(SendMessage);
                    this.Invoke(d, new object[] { str, c, noLog });
                    return;
                }

                str = str.Trim();
                str = str.Replace("\n", " ");
                str = str.Replace("\r", String.Empty);
                str = str.Replace("\t", String.Empty);
                errorMessage.ForeColor = c.HasValue ? c.Value : Color.Black;
                if (str.Length <= 60)
                    errorMessage.Text = str;
                else
                    errorMessage.Text = str.Substring(0, 60) + Environment.NewLine + str.Substring(60, str.Length - 60);

                ClearMessageTimer();

                if (!String.IsNullOrEmpty(str) && !noLog)
                    Log.n(str, c, true);
            }
            catch
            {
                errorMessage.ForeColor = Color.Red;
                errorMessage.Text = "Feil i meldingssystem!" + Environment.NewLine + "Siste feilmelding: " + str;
            }
        }

        private void FormSettingsMacro_Load(object sender, EventArgs e)
        {
            ImportMacroSettings();
        }

        private void FormSettingsMacro_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult != System.Windows.Forms.DialogResult.Cancel)
            {
                if (!ExportMacroSettings() && this.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    SendMessage("Feil format/innstilling.");
                    e.Cancel = true;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!ExportMacroSettings())
                SendMessage("Feil format/innstilling.");
            else
                ImportMacroSettings();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (buttonTestMacro.Text.Equals("Test Makro"))
            {
                buttonTestMacro.Text = "Avbryt";
                StartTestMacro();
            }
            else
                this.abortTest = true;
        }

        private void StartTestMacro()
        {
            try
            {
                if (MessageBox.Show("Advarsel: Alle makro innstillinger blir gjenopprettet til standard "
                    + "og overskriver de gjeldene.\n\nEr du sikker på at du vil fortsette?", "Viktig informasjon", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.Yes)
                {
                    SendMessage("Avbrutt");
                    buttonTestMacro.Text = "Test Makro";
                    return;
                }

                SendMessage("Sjekker innstillinger..");

                if (!statusElguide || !statusEpost || textBoxElguideBrukernavn.Text.Length < 3 || textBoxElguidePassord.Text.Length < 3)
                {
                    buttonTestMacro.Text = "Test Makro";
                    SendMessage("Innstillinger mangler!");
                    return;
                }

                if (this.abortTest)
                {
                    buttonTestMacro.Text = "Test Makro";
                    SendMessage("Avbrutt");
                    this.abortTest = false;
                    return;
                }


                main.appConfig.autoRank = true;
                main.UpdateTimer();
                main.appConfig.AutoStore = true;
                main.UpdateTimerAutoStore();
                main.appConfig.autoQuick = true;
                main.UpdateTimerQuick();
                main.appConfig.AutoService = true;
                main.UpdateTimerService();

                File.WriteAllText(FormMain.macroProgram, Resources.Macro);
                File.WriteAllText(FormMain.macroProgramStore, Resources.MacroStore);
                File.WriteAllText(FormMain.macroProgramQuick, Resources.MacroQuick);
                File.WriteAllText(FormMain.macroProgramService, Resources.MacroService);

                ImportMacroSettings();

                string user = textBoxElguideBrukernavn.Text;
                string pass = textBoxElguidePassord.Text;

                textBoxMacroProgram.Text = textBoxMacroProgram.Text.Replace("brukernavn", user);
                textBoxMacroQuick.Text = textBoxMacroQuick.Text.Replace("brukernavn", user);
                textBoxMacroServiceProgram.Text = textBoxMacroServiceProgram.Text.Replace("brukernavn", user);
                textBoxMacroStore.Text = textBoxMacroStore.Text.Replace("brukernavn", user);

                textBoxMacroProgram.Text = textBoxMacroProgram.Text.Replace("063076096091137154182008082147210214147164196053", pass);
                HashPassword(textBoxMacroProgram);
                textBoxMacroQuick.Text = textBoxMacroQuick.Text.Replace("063076096091137154182008082147210214147164196053", pass);
                HashPassword(textBoxMacroQuick);
                textBoxMacroServiceProgram.Text = textBoxMacroServiceProgram.Text.Replace("063076096091137154182008082147210214147164196053", pass);
                HashPassword(textBoxMacroServiceProgram);
                textBoxMacroStore.Text = textBoxMacroStore.Text.Replace("063076096091137154182008082147210214147164196053", pass);
                HashPassword(textBoxMacroStore);

                if (this.abortTest)
                {
                    buttonTestMacro.Text = "Test Makro";
                    SendMessage("Avbrutt");
                    this.abortTest = false;
                    return;
                }

                ExportMacroSettings();
                main.SaveSettings();

                delayedMacro();
                return;
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
            buttonTestMacro.Text = "Test Makro";
        }

        BackgroundWorker bwMacro = new BackgroundWorker();
        private bool abortTest = false;

        private void delayedMacro()
        {
            try
            {
                for (int b = 0; b < 100; b++)
                {
                    SendMessage("Ser bra ut. Starter makro om " + (((b / 10) * -1) + 10) + " sekunder..", null, true);
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                    if (this.abortTest)
                    {
                        buttonTestMacro.Text = "Test Makro";
                        SendMessage("Avbrutt");
                        this.abortTest = false;
                        return;
                    }
                }

                SendMessage("Importerer transaksjoner med makro..");

                bwMacro.DoWork += new DoWorkEventHandler(bwMacro_DoWork);
                bwMacro.ProgressChanged += new ProgressChangedEventHandler(bwMacro_ProgressChanged);
                bwMacro.WorkerReportsProgress = true;
                bwMacro.WorkerSupportsCancellation = true;
                bwMacro.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwMacro_Completed);

                if (this.abortTest)
                {
                    buttonTestMacro.Text = "Test Makro";
                    SendMessage("Avbrutt");
                    this.abortTest = false;
                    return;
                }

                buttonTestMacro.Text = "Kjører..";
                this.Update();
                this.tabControl1.Enabled = false;
                this.buttonOK.Enabled = false;
                this.buttonSave.Enabled = false;

                bwMacro.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void bwMacro_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                FormMain.autoMode = true;
                SendMessage("Starter test makro..");

                DateTime date = DateTime.Now;
                date = date.AddDays(-2);

                FormMacro form = new FormMacro(main, date, FormMain.macroProgram, 0, true, bwMacro);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.ShowDialog();
                if (form.errorCode != 0)
                {
                    e.Result = "Makro returnerte med feilkode " + form.errorCode + " - Feilmelding: " + form.errorMessage;
                    return;
                }

                SendMessage("Importerer data..");

                var csvFilesToImport = new List<string>() { };
                ImportManager importMng = new ImportManager(main, csvFilesToImport);
                importMng.DoImportTransactions(bwMacro, false);

                if (importMng.returnCode != 0)
                {
                    e.Result = "Importering misslyktes med feilkode " + importMng.returnCode;
                    return;
                }

                SendMessage("Forbereder ranking..");
                main.RetrieveDb(true);

                SendMessage("Lager PDF..");
                string pdf = main.CreatePDF("Full", "", bwMacro);

                List<MailAddress> recip = new List<MailAddress>() { };
                recip.Add(new MailAddress(kgsaEmail.emailDb.Rows[0]["Address"].ToString(), kgsaEmail.emailDb.Rows[0]["Name"].ToString()));

                SendMessage("Sender e-post for mottaker " + kgsaEmail.emailDb.Rows[0]["Address"].ToString() + "..");
                if (!kgsaEmail.InternalSendMail(recip, main.appConfig.epostEmne, main.appConfig.epostBody, new List<string> { pdf }, main.appConfig.epostBrukBcc))
                {
                    e.Result = "Sending av e-post misslyktes";
                    return;
                }

                e.Result = "OK";
            }
            catch (Exception ex)
            {
                SendMessage("Kritisk feil oppstod under kjøring av Test Makro! Se logg for detaljer.", Color.Red);
                Log.Unhandled(ex);
                e.Result = "Exception: " + ex.Message;
            }
        }

        private void bwMacro_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // null
        }

        private void bwMacro_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            this.abortTest = false;
            FormMain.autoMode = false;
            buttonTestMacro.Text = "Test Makro";
            this.Update();
            this.tabControl1.Enabled = true;
            this.buttonOK.Enabled = true;
            this.buttonSave.Enabled = true;
            string result = e.Result.ToString();

            if (e.Error == null && !e.Cancelled && result.Equals("OK"))
            {
                SendMessage("Test Makro fullført", Color.Green);
                labelStatusMacro.Text = "OK";
                labelStatusMacro.ForeColor = Color.Green;
                MessageBox.Show("Test makro fullført uten feil. Makro programmene er nå ferdig klargjort og aktivert."
                    + "\nLegg inn e-poster i Adresseboken til de som skal motta rapportene.\nAnbefales sterkt å la "
                    + "programmet gå en dag før antall mottakere utvides.", "Makro fullført", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                labelStatusMacro.Text = "Feil";
                labelStatusMacro.ForeColor = Color.Red;

                SendMessage("Test Makro feil!", Color.Red);
                MessageBox.Show("Makro avsluttet med feil.\nSjekk innstillinger for Elguide, brukernavn "
                    + "og passord, e-post og generelle program innstillinger.\nSe i loggen for opphavet til feilen.\n\nMelding: "
                    + (string)e.Result, "Makro feil", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            HashPassword(textBoxMacroProgram);
        }

        private void HashPassword(TextBox txtBox)
        {
            try
            {
                string content = txtBox.Text;

                if (content.Contains("Password("))
                {
                    List<string> list = new List<string>(
                           content.Split(new string[] { "\r\n" },
                           StringSplitOptions.RemoveEmptyEntries));

                    int found = 0;
                    for (int i = 0; i < list.Count; i++)
                    {
                        if (list[i].StartsWith("Password"))
                        {
                            string str = Regex.Match(list[i], @"\(([^)]*)\)").Groups[1].Value;
                            if (str.Length > 1 && str.Length < 20 && !Regex.IsMatch(str, @"^\d+$"))
                            {
                                SimpleAES simple = new SimpleAES();
                                string hash = simple.EncryptToString(str);
                                list[i] = "Password(" + hash + ")";
                                found++;
                            }
                        }
                    }
                    if (found > 0)
                        SendMessage(found + " passord kryptert.", Color.Green);
                    else
                        SendMessage("Passord allerede kryptert eller i feil format.");

                    txtBox.Text = string.Join(Environment.NewLine, list.ToArray());
                }
                else
                    SendMessage("Fant ingen linjer med Password()", Color.Red);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            HashPassword(textBoxMacroServiceProgram);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            HashPassword(textBoxMacroQuick);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            HashPassword(textBoxMacroStore);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllText(FormMain.macroProgramStore, Resources.MacroStore);
                textBoxMacroStore.Text = File.ReadAllText(FormMain.macroProgramStore);
                SendMessage("Program 115 reset.", Color.Green);
            }
            catch (Exception ex)
            {
                SendMessage("Feil ved lesing av program 115 makro.", Color.Red);
                Log.Unhandled(ex);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllText(FormMain.macroProgramQuick, Resources.MacroQuick);
                textBoxMacroQuick.Text = File.ReadAllText(FormMain.macroProgramQuick);
                SendMessage("Program 136 reset.", Color.Green);
            }
            catch (Exception ex)
            {
                SendMessage("Feil ved lesing av makro program.", Color.Red);
                Log.Unhandled(ex);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllText(FormMain.macroProgramService, Resources.MacroService);
                textBoxMacroServiceProgram.Text = File.ReadAllText(FormMain.macroProgramService);
                SendMessage("Program 244 reset.", Color.Green);
            }
            catch (Exception ex)
            {
                SendMessage("Feil ved lesing av makro program.", Color.Red);
                Log.Unhandled(ex);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                File.WriteAllText(FormMain.macroProgram, Resources.Macro);
                textBoxMacroProgram.Text = File.ReadAllText(FormMain.macroProgram);
                SendMessage("Program 137 reset.", Color.Green);
            }
            catch (Exception ex)
            {
                SendMessage("Feil ved lesing av makro program.", Color.Red);
                Log.Unhandled(ex);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Browse etter Telnet profil
            try
            {
                var fdlg = new OpenFileDialog();
                fdlg.Title = "Velg Elguide profil";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "Alle filer (*.*)|*.*|Eric's Telnet98 profil|*.etx;*.etn";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    textBoxElguide.Text = fdlg.FileName;
                }
                fdlg.Dispose();
            }
            catch (Exception ex)
            {
                SendMessage("Unntak ved valg av elguide profil.", Color.Red);
                Log.Unhandled(ex);
            }
        }
    }
}
