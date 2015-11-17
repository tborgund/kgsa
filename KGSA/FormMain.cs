using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Printing;
using Microsoft.Win32;

namespace KGSA
{
    delegate void SetTextCallback(string str, Color? c = null, bool logOnly = false, bool statusOnly = false);
    public delegate void DelegateRunServiceList(string s, string t = "");
    public delegate void DelegateRunServiceSpecialList(string status, string filter);

    public partial class FormMain : Form
    {
        #region Variables
        #region Static variables
        public static string version = "v3.0";
        public static string settingsPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\KGSA";
        public static string settingsFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\KGSA\Settings.xml";
        public static string settingsTemp = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\KGSA\Temp";
        public static string settingsWeb = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\KGSA\Web";
        public static string[] staticVaregruppe = new string[154] { "999", "997", "982", "980", "961", "910", "699", "599", "597", "595", "590", "589", "587", "585", "583", "582", "580", "575", "570", "569", "567", "563", "561", "559", "558", "557", "555", "554", "552", "550", "545", "544", "543", "541", "539", "538", "537", "536", "534", "533", "531", "499", "497", "489", "487", "485", "483", "482", "481", "480", "452", "447", "437", "435", "433", "431", "409", "403", "401", "399", "398", "397", "395", "394", "388", "386", "384", "383", "382", "380", "378", "377", "376", "368", "366", "364", "360", "359", "358", "356", "355", "354", "353", "352", "350", "348", "347", "346", "344", "343", "340", "337", "336", "335", "309", "308", "307", "306", "305", "304", "303", "302", "301", "299", "297", "294", "284", "283", "282", "280", "278", "277", "276", "275", "274", "273", "272", "271", "269", "255", "254", "252", "250", "245", "236", "235", "227", "226", "224", "219", "217", "216", "214", "199", "197", "195", "190", "183", "182", "180", "163", "146", "145", "144", "143", "142", "140", "138", "136", "135", "134", "133", "132", "131" };

        public static string databaseVersion = "DbV7";
        public static string databaseName = "Compact" + databaseVersion + ".sdf"; // v7: Ny tblSalg - optimalisert, fjernet varekode tabell.
        public static string SqlConStr = @"Data Source=" + settingsPath + @"\" + databaseName + ";Max Database Size=4091;Max Buffer Size = 1024";
        public static string fileDatabase = settingsPath + @"\" + databaseName;
        // HTML filer
        public static string htmlImport = settingsPath + @"\importdata.html";
        public static string htmlImportService = settingsPath + @"\importdataservice.html";
        public static string htmlImportStore = settingsPath + @"\importstore.html";
        public static string htmlSetupBudget = settingsPath + @"\setupbudget.html";
        public static string htmlLoading = settingsPath + @"\loading.html";
        public static string htmlStopped = settingsPath + @"\stopped.html";
        public static string htmlError = settingsPath + @"\error.html";
        public static string htmlRankingButikk = settingsPath + @"\rankingButikk.html";
        public static string htmlRankingKnowHow = settingsPath + @"\rankingKnowHow.html";
        public static string htmlRankingData = settingsPath + @"\rankingData.html";
        public static string htmlRankingAudioVideo = settingsPath + @"\rankingAudioVideo.html";
        public static string htmlRankingTele = settingsPath + @"\rankingTele.html";
        public static string htmlRankingOversikt = settingsPath + @"\rankingOversikt.html";
        public static string htmlRankingBudget = settingsPath + @"\rankingBudsjett.html";
        public static string htmlRankingToppselgere = settingsPath + @"\rankingToppselgere.html";
        public static string htmlRankingLister = settingsPath + @"\rankingLister.html";
        public static string htmlRankingVinn = settingsPath + @"\rankingVinnprodukter.html";
        public static string htmlRankingVinnSelger = settingsPath + @"\rankingVinnprodukterSelger.html";
        public static string htmlRankingQuick = settingsPath + @"\rankingQuick.html";

        public static string htmlAvdTjenester = settingsPath + @"\rankingAvdTjenester.html";
        public static string htmlAvdSnittpriser = settingsPath + @"\rankingAvdSnittpriser.html";

        public static string htmlBudgetMdaFile = "budsjettMda.html";
        public static string htmlBudgetAudioVideoFile = "budsjettAudioVideo.html";
        public static string htmlBudgetSdaFile = "budsjettSda.html";
        public static string htmlBudgetTeleFile = "budsjettTele.html";
        public static string htmlBudgetDataFile = "budsjettData.html";
        public static string htmlBudgetCrossFile = "budsjettCross.html";
        public static string htmlBudgetKasseFile = "budsjettKasse.html";
        public static string htmlBudgetAftersalesFile = "budsjettAftersales.html";
        public static string htmlBudgetMdasdaFile = "budsjettMdaSda.html";
        public static string htmlBudgetButikkFile = "budsjettButikk.html";
        public static string htmlBudgetDailyFile = "budsjettDaglig.html";
        public static string htmlBudgetAllSalesFile = "budsjettAlleSelgere.html";

        public static string htmlBudgetMda = settingsPath + "\\" + htmlBudgetMdaFile;
        public static string htmlBudgetAudioVideo = settingsPath + "\\" + htmlBudgetAudioVideoFile;
        public static string htmlBudgetSda = settingsPath + "\\" + htmlBudgetSdaFile;
        public static string htmlBudgetTele = settingsPath + "\\" + htmlBudgetTeleFile;
        public static string htmlBudgetData = settingsPath + "\\" + htmlBudgetDataFile;
        public static string htmlBudgetCross = settingsPath + "\\" + htmlBudgetCrossFile;
        public static string htmlBudgetKasse = settingsPath + "\\" + htmlBudgetKasseFile;
        public static string htmlBudgetAftersales = settingsPath + "\\" + htmlBudgetAftersalesFile;
        public static string htmlBudgetMdasda = settingsPath + "\\" + htmlBudgetMdasdaFile;
        public static string htmlBudgetButikk = settingsPath + "\\" + htmlBudgetButikkFile;
        public static string htmlBudgetDaily = settingsPath + "\\" + htmlBudgetDailyFile;
        public static string htmlBudgetAllSales = settingsPath + "\\" + htmlBudgetAllSalesFile;

        public static string htmlRapport = settingsPath + @"\rankingRapport.html";
        public static string htmlPeriode = settingsPath + @"\rankingPeriode.html";
        public static string htmlGraf = settingsPath + @"\rankingGraf.html";
        public static string htmlServiceOversikt = settingsPath + @"\serviceOversikt.html";
        public static string htmlServiceDetails = settingsPath + @"\serviceDetails.html";
        public static string htmlServiceList = settingsPath + @"\serviceList.html";

        public static string htmlStoreObsolete = settingsPath + @"\storeLagerstatus.html";
        public static string htmlStoreObsoleteList = settingsPath + @"\storeObsoleteList.html";        
        public static string htmlStorePrisguide = settingsPath + @"\storePrisguide.html";
        public static string htmlStorePrisguideOverview = settingsPath + @"\storePrisguideOverview.html";
        public static string htmlStoreWeekly = settingsPath + @"\storeWeekly.html";
        public static string htmlStoreWeeklyOverview = settingsPath + @"\storeWeeklyOverview.html";
        public static string htmlStoreObsoleteImports = settingsPath + @"\storeObsoleteImports.html";
        // Tredjeparts program
        public static string fileVarekoder = settingsPath + @"\varekoder.txt";
        public static string filePDFwkhtmltopdf = Application.StartupPath + @"\wkhtmltopdf.exe";

        public static string macroProgram = settingsPath + @"\macro.txt";
        public static string macroProgramQuick = settingsPath + @"\macroQuick.txt";
        public static string macroProgramService = settingsPath + @"\macroService.txt";
        public static string macroProgramStore = settingsPath + @"\macroStore.txt";
        public static string jsJqueryTablesorter = settingsPath + @"\jquery.tablesorter.js";
        public static string jsJqueryMetadata = settingsPath + @"\jquery.metadata.js";
        public static string jsJquery = settingsPath + @"\jquery.js";
        public static Color[] favColors = new Color[] { Color.Green, Color.DarkOrange, Color.DarkKhaki, Color.Magenta, Color.Peru,
            Color.LightBlue, Color.Blue, Color.Red, Color.Orange, Color.Yellow, Color.Purple, Color.Pink, Color.Violet, Color.Lime,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray,
            Color.Gray, Color.Gray, Color.Gray, Color.Gray, Color.Gray}; // 64 farger
        #endregion
        #region BackgroundWorkers
        BackgroundWorker bwGraph = new BackgroundWorker();
        BackgroundWorker bwPopulateSk = new BackgroundWorker();
        BackgroundWorker bwMakeScreens = new BackgroundWorker();
        BackgroundWorker bwQuickAuto = new BackgroundWorker();
        BackgroundWorker bwService = new BackgroundWorker();
        BackgroundWorker bwServiceReport = new BackgroundWorker();
        BackgroundWorker bwUpdateTopGraph = new BackgroundWorker();
        BackgroundWorker bwUpdateBigGraph = new BackgroundWorker();
        BackgroundWorker bwAutoImportService = new BackgroundWorker();
        BackgroundWorker bwImportService = new BackgroundWorker();
        BackgroundWorker bwRanking = new BackgroundWorker();
        BackgroundWorker bwBudget = new BackgroundWorker();
        BackgroundWorker bwStore = new BackgroundWorker();
        BackgroundWorker bwReport = new BackgroundWorker();
        BackgroundWorker bwSendEmail = new BackgroundWorker();
        BackgroundWorker bwImport = new BackgroundWorker();
        BackgroundWorker bwPDF = new BackgroundWorker();
        BackgroundWorker bwBudgetPDF = new BackgroundWorker();
        BackgroundWorker bwOpenPDF = new BackgroundWorker();
        BackgroundWorker bwAutoRanking = new BackgroundWorker();
        BackgroundWorker bwMacroRanking = new BackgroundWorker();
        BackgroundWorker bwAutoStore = new BackgroundWorker();
        BackgroundWorker bwImportObsolete = new BackgroundWorker();
        BackgroundWorker bwCreateHtml = new BackgroundWorker();
        public BackgroundWorker bwHentAvdelinger = new BackgroundWorker();
        BackgroundWorker bwVinnSelger = new BackgroundWorker();
        public BackgroundWorker worker;
        #endregion
        public readonly Timer timerAutoStore = new Timer();
        public DateTime timerNextRunAutoStore = rangeMin;
        public readonly Timer timerAutoService = new Timer();
        public DateTime timerNextRunService = rangeMin;
        public readonly Timer timerAutoRanking = new Timer();
        public DateTime timerNextRun = rangeMin;
        public readonly Timer timerAutoQuick = new Timer();
        public DateTime timerNextRunQuick = rangeMin;
        public TimeWatch timewatch = new TimeWatch();
        private bool Loaded;
        public static CultureInfo norway = new CultureInfo("nb-NO");

        private DateTime chkPicker;
        private DateTime chkBudgetPicker;
        private DateTime chkServicePicker;
        private DateTime chkStorePicker;
        public static DateTime rangeMin = new DateTime(2000, 1, 1, 0, 0, 0);
        public static DateTime rangeMax = DateTime.Now.AddYears(1);
        public static DateTime highlightDate = rangeMin;
        public static bool appManagerIsBusy = false;
        private readonly Timer timerMsgClear = new Timer();

        public AppSettings appConfig = new AppSettings();
        private int lagretAvdeling = 0;
        public Avdeling avdeling = new Avdeling();
        public FormProcessing processing;
        public static string filenamePDF;
        private string lastRightClickValue = "";
        public static bool autoMode;
        public static string[] args;
        public static List<string> Favoritter = new List<string> { };
        public Service service = new Service();
        public Obsolete obsolete = new Obsolete();
        public TopGraph topgraph;
        public GraphClass gc;
        public SalesCodes salesCodes;
        private bool forceShutdown = false;
        private static bool newInstall = false;
        public DataTable sqlceCurrentMonth;
        public Vinnprodukt vinnprodukt;
        public BudgetObj budget;
        public OpenXML openXml;
        public Database database;
        public SqlCeConnection connection = new SqlCeConnection(SqlConStr);
        BluetoothServer blueServer = null;
        public KgsaTools tools;
        public DataTable tableMacroQuick = null;

        #endregion
        public FormMain(string[] arrayArgs)
        {
            var splash = new FormSplash();
            splash.Show();
            splash.Update();
            args = arrayArgs;
            StartupCheck();
            splash.ProgressMesssage("Initialiserer..");
            SetBrowserFeatureControl();
            InitializeComponent();
            InitializeMyComponents();
            LoadSettings();
            splash.ProgressMesssage("Starter tjenester..");
            StartServices();
            GenerateFavorities();
            lagretAvdeling = appConfig.Avdeling;
            splash.ProgressMesssage("Åpner databasen..");
            RetrieveDb();
            RetrieveDbStore();
            RetrieveDbService();
            splash.ProgressMesssage("Leser fra database..");
            ReloadStore();
            ReloadBudget();
            Reload();
            ReloadService();
            splash.ProgressMesssage("Registrerer..");
            UpdateUi();
            splash.Hide();
            splash.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (appConfig.WindowMax)
                this.WindowState = FormWindowState.Maximized;
            else
            {
                if (appConfig.WindowLocationX > 0 && appConfig.WindowLocationY > 0)
                    this.Location = new Point(appConfig.WindowLocationX, appConfig.WindowLocationY);
                this.Size = new Size(appConfig.WindowSizeX, appConfig.WindowSizeY);
                this.WindowState = FormWindowState.Normal;
            }

            listboxContextMenu = new ContextMenuStrip();
            listboxContextMenu.Opening += new CancelEventHandler(listboxContextMenu_Opening);
            listBoxSk.ContextMenuStrip = listboxContextMenu;

            if (appConfig.Avdeling < 1000)
                NewInstallation();
            if (appConfig.Avdeling < 1000) // Sjekk for manglende innstillinger..
                Log.Alert("Viktig: Avdeling er IKKE valgt.\n\nGå til innstillinger og legg inn din avdeling!", "Manglende innstillinger", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Log.n("KGSA " + version + " (Build: " + RetrieveLinkerTimestamp().ToShortDateString() + ")", Color.Black, true);
            Focus();
            Loaded = true; // Forhindrer controls on-change events å fyre av under lasting av programmet.
            if (args.Length > 0)
            {
                string filename = args[0];
                try
                {
                    if (filename.Length > 6 && File.Exists(filename))
                    {
                        if (filename.EndsWith("iserv.csv"))
                        {
                            Log.n("Angitt CSV (" + filename + ") er av type 'Elguide Service Logg'. Starter importering..");
                            bwImportService.RunWorkerAsync(filename);
                            processing.SetVisible = true;
                            processing.SetBackgroundWorker = bwImportService;
                            processing.SetText = "Importerer service CSV..";
                        }
                        else if (filename.EndsWith("irank.csv"))
                        {
                            Log.n("Angitt CSV (" + filename + ") er av type 'Elguide Ranking Transaksjoner'. Starter importering..");
                            csvFilesToImport.Clear();
                            csvFilesToImport.Add(args[0]);
                            RunImport(true);
                        }
                        else
                        {

                            var velgcsv = new VelgTypeCSV();
                            velgcsv.ShowDialog();

                            if (velgcsv.DialogResult == System.Windows.Forms.DialogResult.No)
                            {
                                Log.n("Angitt CSV (" + filename + ") er av type 'Elguide Service Logg'. Starter importering..");
                                bwImportService.RunWorkerAsync(filename);
                                processing.SetVisible = true;
                                processing.SetBackgroundWorker = bwImportService;
                                processing.SetText = "Importerer service CSV..";
                            }
                            else if (velgcsv.DialogResult == System.Windows.Forms.DialogResult.Yes)
                            {
                                Log.n("Angitt CSV (" + filename + ") er av type 'Elguide Ranking Transaksjoner'. Starter importering..");
                                csvFilesToImport.Clear();
                                csvFilesToImport.Add(args[0]);
                                RunImport(true);
                            }
                            else
                                Log.n("Angitt CSV (" + filename + ") ble ikke gjenkjent av KGSA.", Color.Red);
                        }
                    }
                    else
                        Log.n("Ukjent argument: " + filename + ".", Color.Red);
                }
                catch(Exception ex)
                {
                    Log.Unhandled(ex);
                }
            }

            if (RetrieveLinkerTimestamp() < DateTime.Now.AddYears(-2))
            {
                MessageBox.Show("Denne versjonen av KGSA er utdatert. Anbefaler sterkt å oppdatere.", "KGSA - Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (appConfig.WindowExitToTray && !forceShutdown)
            {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.Visible = true;
                notifyIcon1.BalloonTipText = "Kjører fortsatt i bakgrunnen." + Environment.NewLine + "Fil -> Avslutt for å avslutte programmet.";
                notifyIcon1.ShowBalloonTip(3000);
                this.ShowInTaskbar = false;
                e.Cancel = true;
                return;
            }

            if (!bwImport.IsBusy && !bwImportService.IsBusy)
            {
                // Copy window location to app settings
                appConfig.WindowLocationX = this.Location.X;
                appConfig.WindowLocationY = this.Location.Y;

                // Copy window size to app settings
                if (this.WindowState == FormWindowState.Normal)
                {
                    appConfig.WindowSizeX = this.Size.Width;
                    appConfig.WindowSizeY = this.Size.Height;
                }
                else
                {
                    appConfig.WindowSizeX = this.RestoreBounds.Size.Width;
                    appConfig.WindowSizeY = this.RestoreBounds.Size.Height;
                }

                if (this.WindowState == FormWindowState.Maximized)
                    appConfig.WindowMax = true;
                else
                    appConfig.WindowMax = false;

                if (lagretAvdeling != appConfig.Avdeling) // Hvis favoritt-velger er benyttet, lagre avdelingsnummer som var sist satt i innstillinger.
                    appConfig.Avdeling = lagretAvdeling;

                appConfig.savedTab = tabControlMain.SelectedTab.Text;

                if (arrayDbAvd != null)
                    if (arrayDbAvd.Length > 0)
                        appConfig.avdelingerListAlle = arrayDbAvd.ToList();

                openXml.SaveDatabase();

                if (blueServer != null)
                    if (blueServer.IsOnline())
                        blueServer.StopServer();

                SaveSettings();
                notifyIcon1.Dispose();
                Log.LogAdded -= new EventHandler(LogMessage_LogAdded);
            }
            else
            {
                forceShutdown = false;
                e.Cancel = true;
            }
        }

        private void InitializeMyComponents()
        {
            database = new Database(this);
            vinnprodukt = new Vinnprodukt(this);
            openXml = new OpenXML(this);
            processing = new FormProcessing(this);
            salesCodes = new SalesCodes(this);
            tools = new KgsaTools(this);

            worker = new BackgroundWorker();

            webService.ObjectForScripting = new ScriptInterface();
            UpdateServicePage.OnBrowseServicePage += UpdateServicePage_OnRun;
            Log.LogAdded += new EventHandler(LogMessage_LogAdded);

            timerAutoRanking.Tick += timer_Tick;
            timerAutoQuick.Tick += timer_TickQuick;
            timerAutoService.Tick += timerService_Tick;
            timerAutoStore.Tick += timerStore_Tick;
            timerMsgClear.Tick += timer;

            bwSendEmail.DoWork += new DoWorkEventHandler(bwSendEmail_DoWork);
            bwSendEmail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwSendEmail_RunWorkerCompleted);

            bwPopulateSk.DoWork += new DoWorkEventHandler(bwPopulateSk_DoWork);
            bwPopulateSk.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPopulateSk_Completed);

            bwImport.DoWork += new DoWorkEventHandler(bwImport_DoWork);
            bwImport.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwImport.WorkerReportsProgress = true;
            bwImport.WorkerSupportsCancellation = true;
            bwImport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwImport_Completed);

            bwImportObsolete.DoWork += new DoWorkEventHandler(bwImportObsolete_DoWork);
            bwImportObsolete.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwImportObsolete.WorkerReportsProgress = true;
            bwImportObsolete.WorkerSupportsCancellation = true;
            bwImportObsolete.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwImportObsolete_Completed);

            bwPDF.DoWork += new DoWorkEventHandler(bwPDF_DoWork);
            bwPDF.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwPDF.WorkerReportsProgress = true;
            bwPDF.WorkerSupportsCancellation = true;
            bwPDF.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPDF_Completed);

            bwBudgetPDF.DoWork += new DoWorkEventHandler(bwBudgetPDF_DoWork);
            bwBudgetPDF.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwBudgetPDF.WorkerReportsProgress = true;
            bwBudgetPDF.WorkerSupportsCancellation = true;
            bwBudgetPDF.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwPDF_Completed);

            bwOpenPDF.DoWork += new DoWorkEventHandler(bwOpenPDF_DoWork);
            bwOpenPDF.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwOpenPDF.WorkerReportsProgress = true;
            bwOpenPDF.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwOpenPDF_Completed);

            bwAutoRanking.DoWork += new DoWorkEventHandler(bwAutoRanking_DoWork);
            bwAutoRanking.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwAutoRanking.WorkerReportsProgress = true;
            bwAutoRanking.WorkerSupportsCancellation = true;
            bwAutoRanking.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwAutoRanking_Completed);

            bwMacroRanking.DoWork += new DoWorkEventHandler(bwMacroRanking_DoWork);
            bwMacroRanking.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwMacroRanking.WorkerReportsProgress = true;
            bwMacroRanking.WorkerSupportsCancellation = true;
            bwMacroRanking.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwMacroRanking_Completed);

            bwAutoStore.DoWork += new DoWorkEventHandler(bwAutoStore_DoWork);
            bwAutoStore.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwAutoStore.WorkerReportsProgress = true;
            bwAutoStore.WorkerSupportsCancellation = true;
            bwAutoStore.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwAutoStore_Completed);

            bwRanking.DoWork += new DoWorkEventHandler(bwRanking_DoWork);
            bwRanking.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwRanking.WorkerReportsProgress = true;
            bwRanking.WorkerSupportsCancellation = true;
            bwRanking.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwRanking_Completed);

            bwVinnSelger.DoWork += new DoWorkEventHandler(bwVinnSelger_DoWork);
            bwVinnSelger.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwVinnSelger.WorkerReportsProgress = true;
            bwVinnSelger.WorkerSupportsCancellation = true;
            bwVinnSelger.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwRanking_Completed);

            bwBudget.DoWork += new DoWorkEventHandler(bwBudget_DoWork);
            bwBudget.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwBudget.WorkerReportsProgress = true;
            bwBudget.WorkerSupportsCancellation = true;
            bwBudget.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwBudget_Completed);

            bwStore.DoWork += new DoWorkEventHandler(bwStore_DoWork);
            bwStore.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwStore.WorkerReportsProgress = true;
            bwStore.WorkerSupportsCancellation = true;
            bwStore.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwStore_Completed);

            bwReport.DoWork += new DoWorkEventHandler(bwReport_DoWork);
            bwReport.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwReport.WorkerReportsProgress = true;
            bwReport.WorkerSupportsCancellation = true;
            bwReport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwReport_Completed);

            bwQuickAuto.DoWork += new DoWorkEventHandler(bwQuickAuto_DoWork);
            bwQuickAuto.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwQuickAuto.WorkerReportsProgress = true;
            bwQuickAuto.WorkerSupportsCancellation = true;
            bwQuickAuto.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwQuickAuto_Completed);

            bwAutoImportService.DoWork += new DoWorkEventHandler(bwAutoImportService_DoWork);
            bwAutoImportService.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwAutoImportService.WorkerReportsProgress = true;
            bwAutoImportService.WorkerSupportsCancellation = true;
            bwAutoImportService.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwAutoImportService_Completed);

            bwImportService.DoWork += new DoWorkEventHandler(bwImportService_DoWork);
            bwImportService.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwImportService.WorkerReportsProgress = true;
            bwImportService.WorkerSupportsCancellation = true;
            bwImportService.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwImportService_Completed);

            bwGraph.DoWork += new DoWorkEventHandler(bwGraph_DoWork);
            bwGraph.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwGraph.WorkerReportsProgress = true;
            bwGraph.WorkerSupportsCancellation = true;
            bwGraph.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwGraph_Completed);

            bwService.DoWork += new DoWorkEventHandler(bwService_DoWork);
            bwService.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwServiceGraph_Completed);

            bwServiceReport.DoWork += new DoWorkEventHandler(bwServiceReport_DoWork);
            bwServiceReport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwServiceReport_Completed);

            bwUpdateTopGraph.DoWork += new DoWorkEventHandler(bwUpdateTopGraph_DoWork);
            bwUpdateTopGraph.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUpdateTopGraph_Completed);

            bwUpdateBigGraph.DoWork += new DoWorkEventHandler(bwUpdateBigGraph_DoWork);
            bwUpdateBigGraph.ProgressChanged += new ProgressChangedEventHandler(bwProgressCustom_ProgressChanged);
            bwUpdateBigGraph.WorkerSupportsCancellation = true;
            bwUpdateBigGraph.WorkerReportsProgress = true;
            bwUpdateBigGraph.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUpdateBigGraph_Completed);

            bwCreateHtml.DoWork += new DoWorkEventHandler(bwCreateHtml_DoWork);
            bwCreateHtml.ProgressChanged += new ProgressChangedEventHandler(bwProgressReport_ProgressChanged);
            bwCreateHtml.WorkerReportsProgress = true;
            bwCreateHtml.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwCreateHtml_Completed);

            bwHentAvdelinger.DoWork += new DoWorkEventHandler(bwHentAvdelinger_DoWork);
            bwHentAvdelinger.WorkerSupportsCancellation = true;
            bwHentAvdelinger.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwHentAvdelinger_Completed);

        }

        private void SetBrowserFeatureControl()
        {
            // http://msdn.microsoft.com/en-us/library/ee330720(v=vs.85).aspx

            // FeatureControl settings are per-process
            var fileName = System.IO.Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);

            // make the control is not running inside Visual Studio Designer
            if (String.Compare(fileName, "devenv.exe", true) == 0 || String.Compare(fileName, "XDesProc.exe", true) == 0)
                return;

            SetBrowserFeatureControlKey("FEATURE_BROWSER_EMULATION", fileName, GetBrowserEmulationMode());
            // Webpages containing standards-based !DOCTYPE directives are displayed in IE10 Standards mode.
            SetBrowserFeatureControlKey("FEATURE_AJAX_CONNECTIONEVENTS", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_MANAGE_SCRIPT_CIRCULAR_REFS", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_DOMSTORAGE ", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_GPU_RENDERING ", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_IVIEWOBJECTDRAW_DMLT9_WITH_GDI  ", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_DISABLE_LEGACY_COMPRESSION", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_LOCALMACHINE_LOCKDOWN", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_BLOCK_LMZ_OBJECT", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_BLOCK_LMZ_SCRIPT", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_DISABLE_NAVIGATION_SOUNDS", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_SCRIPTURL_MITIGATION", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_SPELLCHECKING", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_STATUS_BAR_THROTTLING", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_TABBED_BROWSING", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_VALIDATE_NAVIGATE_URL", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_WEBOC_DOCUMENT_ZOOM", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_WEBOC_POPUPMANAGEMENT", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_WEBOC_MOVESIZECHILD", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_ADDON_MANAGEMENT", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_WEBSOCKET", fileName, 1);
            SetBrowserFeatureControlKey("FEATURE_WINDOW_RESTRICTIONS ", fileName, 0);
            SetBrowserFeatureControlKey("FEATURE_XMLHTTP", fileName, 1);
        }

        private UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 9;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            UInt32 mode = 11000; // Internet Explorer 11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 Standards mode. Default value for Internet Explorer 11.
            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. Default value for applications hosting the WebBrowser Control.
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. Default value for Internet Explorer 8
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode. Default value for Internet Explorer 9.
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10. Webpages containing standards-based !DOCTYPE directives are displayed in IE10 mode. Default value for Internet Explorer 10.
                    break;
                default:
                    // use IE11 mode by default
                    break;
            }

            return mode;
        }

        private void SetBrowserFeatureControlKey(string feature, string appName, uint value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(
                String.Concat(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\", feature),
                RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                key.SetValue(appName, (UInt32)value, RegistryValueKind.DWord);
            }
        }

        void UpdateServicePage_OnRun(string s, string t = "")
        {
            if (s == "ServiceList")
                RunServiceList(t);
            else if (s == "ServiceOversikt")
                RunServiceOversikt();
        }

        private void timer(object sender, EventArgs e)
        {
            ClearMessageTimerStop();
        }

        private void ClearMessageTimer()
        {
            timerMsgClear.Stop();
            timerMsgClear.Interval = 30 * 1000;
            timerMsgClear.Enabled = true;
            timerMsgClear.Start();
        }

        private void ClearMessageTimerStop()
        {
            if (timerMsgClear.Enabled)
            {
                Log.Status("");
                timerMsgClear.Enabled = false;
                timerMsgClear.Stop();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                if (bwRanking.IsBusy)
                {
                    stopRanking = true;
                    Log.n("Avbryter ranking..", Color.Red);
                    return true;
                }
                if (bwUpdateBigGraph.IsBusy)
                {
                    _graphReqStop = true;
                    Log.n("Avbryter graf..", Color.Red);
                    return true;
                }
                return false;
            }
            if (keyData == (Keys.Control | Keys.P))
            {
                string curTab = readCurrentTab();
                if (curTab == "Ranking")
                    webRanking.ShowPrintDialog();
                else if (curTab == "Service")
                    webService.ShowPrintDialog();
                else if (curTab == "Store")
                    webStore.ShowPrintDialog();
                else
                    Log.n("Gjeldene vindu kan ikke skrives ut.");
                return true;
            }
            if (keyData == (Keys.Shift | Keys.P))
            {
                RunOpenPDF();
                return true;
            }
            if (keyData == Keys.F5)
            {
                if (!IsBusy())
                {
                    buttonOppdater.PerformClick();
                    return true;
                }
                return false;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void FavorittClick(object sender, EventArgs e)
        {
            try
            {
                int avd = Convert.ToInt32(sender.ToString().Substring(0, 4));
                if (avd > 0 && avd != appConfig.Avdeling && !IsBusy())
                {
                    processing.SetVisible = true;
                    appConfig.Avdeling = avd;
                    SaveSettings();
                    processing.SetValue = 10;
                    RetrieveDbService();
                    RetrieveDbStore();
                    RetrieveDb();
                    processing.SetValue = 50;
                    ReloadService();
                    ReloadStore(true);
                    Reload(true);
                    processing.SetValue = 80;
                    UpdateUi();
                    tabControlMain_SelectedIndexChanged(sender, e);
                    processing.SetValue = 100;
                    processing.HideDelayed();
                    this.Activate();
                }                    
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void FavorittClickAdd(object sender, EventArgs e)
        {
            try
            {
                int avd = Convert.ToInt32(sender.ToString().Substring(0, 4));
                if (avd > 1000 && !appConfig.favAvdeling.Contains(avd.ToString()))
                {
                    Log.n("Legger til avdeling..", null, false, true);
                    if (appConfig.favAvdeling.Length > 3) // Den er ikke tom
                        appConfig.favAvdeling += "," + avd;
                    else
                        appConfig.favAvdeling = avd.ToString();
                    if (!appConfig.favVis)
                        appConfig.favVis = true;
                    ClearHash();
                    SaveSettings();
                    UpdateUi();
                    Log.n("Lagt til " + avdeling.Get(avd) + " i favoritt avdelinger.", Color.Green);
                }
                else
                    Log.n("Avdelingen finnes i favoritt listen fra før.", Color.Red);
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil: Kunne ikke legge til avdeling.", Color.Red);
            }
        }

        private void omToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var aboutForm = new FormAbout();
            aboutForm.StartPosition = FormStartPosition.CenterParent;
            aboutForm.ShowDialog();
            aboutForm.Dispose();
        }

        private void avsluttToolStripMenuItem_Click(object sender, EventArgs e)
        {
            forceShutdown = true;
            Application.Exit();
        }

        private void rankingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
            if (!IsBusy(true))
                RunRanking("AudioVideo");
        }

        private void kalkulatorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageGrafikk;
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageTrans;
        }

        private void loggToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageLog;
        }

        private void print(object sender, EventArgs e)
        {
            string curTab = readCurrentTab();
            if (curTab == "Ranking")
                webRanking.ShowPrintDialog();
            else if (curTab == "Budget")
                webBudget.ShowPrintDialog();
            else if (curTab == "Service")
                webService.ShowPrintDialog();
            else if (curTab == "Store")
                webStore.ShowPrintDialog();
            else if (curTab == "Log")
            {
                PrintDialog printDialog = new PrintDialog();
                PrintDocument documentToPrint = new PrintDocument();
                printDialog.Document = documentToPrint;
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    StringReader reader = new StringReader(richLog.Text);
                    documentToPrint.PrintPage += new PrintPageEventHandler(DocumentToPrint_PrintPage);
                    documentToPrint.Print();
                }
            }
            else
                Log.n("Gjeldene vindu kan ikke skrives ut.");
        }

        private void DocumentToPrint_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            StringReader reader = new StringReader(richLog.Text);
            float LinesPerPage = 0;
            float YPosition = 0;
            int Count = 0;
            float LeftMargin = e.MarginBounds.Left;
            float TopMargin = e.MarginBounds.Top;
            string Line = null;
            Font PrintFont = this.richLog.Font;
            SolidBrush PrintBrush = new SolidBrush(Color.Black);

            LinesPerPage = e.MarginBounds.Height / PrintFont.GetHeight(e.Graphics);

            while (Count < LinesPerPage && ((Line = reader.ReadLine()) != null))
            {
                YPosition = TopMargin + (Count * PrintFont.GetHeight(e.Graphics));
                e.Graphics.DrawString(Line, PrintFont, PrintBrush, LeftMargin, YPosition, new StringFormat());
                Count++;
            }

            if (Line != null)
            {
                e.HasMorePages = true;
            }
            else
            {
                e.HasMorePages = false;
            }
            PrintBrush.Dispose();
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            try
            {
                if (!toolStripMenuItemVisHist.Checked)
                {
                    panel4.Visible = true;
                    toolStripMenuItemVisHist.Checked = true;
                    appConfig.histogramVis = true;
                    RunTopGraphUpdate();
                }
                else
                {
                    panel4.Visible = false;
                    toolStripMenuItemVisHist.Checked = false;
                    appConfig.histogramVis = false;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }


        private void timer_Tick(object sender, EventArgs e)
        {
            UpdateTimer();
        }

        private void timer_TickQuick(object sender, EventArgs e)
        {
            UpdateTimerQuick();
        }

        private void timerService_Tick(object sender, EventArgs e)
        {
            UpdateTimerService();
        }

        private void timerStore_Tick(object sender, EventArgs e)
        {
            UpdateTimerAutoStore();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!panel4.Visible)
            {
                panel4.Visible = true;
                toolStripMenuItemVisHist.Checked = true;
                appConfig.histogramVis = true;
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
                appConfig.histogramVis = false;
            }
        }

        private void toolMenuSettings_Click(object sender, EventArgs e)
        {
            OpenSettings();
        }

        private void toolMenuSendrank_Click(object sender, EventArgs e)
        {
            if (!IsBusy() && !EmptyDatabase())
                OpenSendEmail();
        }

        private void toolStripMenuItem7_Click_1(object sender, EventArgs e)
        {
            string curTab = readCurrentTab();
            if (curTab == "Ranking")
            {
                webRanking.ShowPrintPreviewDialog();
                return;
            }
            else if (curTab == "Budget")
            {
                webBudget.ShowPrintPreviewDialog();
                return;
            }
            else if (curTab == "Service")
            {
                webService.ShowPrintPreviewDialog();
                return;
            }
            else if (curTab == "Store")
            {
                webStore.ShowPrintPreviewDialog();
                return;
            }
            
            Log.n("Gjeldene side kan ikke skrives ut.");
        }

        private void toolMenuPagesetup_Click(object sender, EventArgs e)
        {
            string curTab = readCurrentTab();
            if (curTab == "Ranking")
            { 
                webRanking.ShowPageSetupDialog();
                return;
            }
            if (curTab == "Avdelinger")
            {
                webRanking.ShowPageSetupDialog();
                return;
            }
            else if (curTab == "Budget")
            {
                webBudget.ShowPageSetupDialog();
                return;
            }
            else if (curTab == "Service")
            {
                webService.ShowPageSetupDialog();
                return;
            }
            else if (curTab == "Store")
            {
                webStore.ShowPageSetupDialog();
                return;
            }
            
            Log.n("Gjeldene side kan ikke skrives ut.");
        }

        private void graphPanelTop_Paint(object sender, PaintEventArgs e)
        {
            PaintHistTopNew();
        }

        private void graphPanelTop_Resize(object sender, EventArgs e)
        {
            graphPanelTop.Invalidate();
        }

        private void graphPanelTop_MouseClick(object sender, MouseEventArgs e)
        {
            clickTopGraph(e);
        }

        private List<string> csvFilesToImport = new List<string> ();

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
            if (!IsBusy(true))
                RunRanking("Data");
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
            if (!IsBusy(true))
                RunRanking("Tele");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            processing.SetVisible = true;
            processing.SetText = "Søker..";
            SearchDB();
            processing.HideDelayed();
            this.Activate();
            Log.Status("Ferdig.");
        }

        private void btokrToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (appConfig.kolInntjen)
            {
                appConfig.kolInntjen = false;
                btokrToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.kolInntjen = true;
                btokrToolStripMenuItem.Checked = true;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void provisjonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (appConfig.kolProv)
            {
                appConfig.kolProv = false;
                provisjonToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.kolProv = true;
                provisjonToolStripMenuItem.Checked = true;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void varekToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!appConfig.kolVarekoder)
            {
                appConfig.kolVarekoder = true;
                varekToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.kolVarekoder = false;
                varekToolStripMenuItem.Checked = true;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void omsetningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (appConfig.kolSalgspris)
            {
                appConfig.kolSalgspris = false;
                omsetningToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.kolSalgspris = true;
                omsetningToolStripMenuItem.Checked = true;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void eksporterDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            exportDatabase();
            processing.HideDelayed();
            this.Activate();
        }

        private void importerDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            importDatabase();
            processing.HideDelayed();
            this.Activate();
        }

        private void lesMegToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Process lesmeg = new Process();
                lesmeg.StartInfo.FileName = Application.StartupPath + @"\Lesmeg.rtf";
                lesmeg.Start();
            }
            catch
            {
                Log.n("Kunne ikke åpne Lesmeg.rtf. Dokumentet mangler eller systemet har ingen programmer som kan åpne den (!).", Color.Red);
            }
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                savePDF(true);
        }

        private void pickerDBTil_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (pickerDBFra.Value > pickerDBTil.Value && Loaded)
                    pickerDBFra.Value = pickerDBTil.Value;
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void pickerButikk_DropDown(object sender, EventArgs e)
        {
            chkPicker = pickerRankingDate.Value;
        }

        private void pickerButikk_CloseUp(object sender, EventArgs e)
        {
            if (chkPicker != pickerRankingDate.Value)
                moveDate(0, true);
        }

        private void webHTML_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            NavigateWeb(e);
        }

        private void buttonToppselgere_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Toppselgere");
        }

        private void buttonOversikt_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Oversikt");
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Butikk");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Data");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("AudioVideo");
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Tele");
        }

        private void button23_Click(object sender, EventArgs e)
        {
            if (!bwRanking.IsBusy && !bwReport.IsBusy)
                UpdateRank();
            else if (buttonOppdater.Text == "Stop" && (bwRanking.IsBusy || bwReport.IsBusy))
                stopRanking = true;
            else if (buttonOppdater.Text == "Stop" && (!bwRanking.IsBusy && !bwReport.IsBusy))
                ProgressStop();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            velgCSV();
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            RunImport();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (IsBusy())
                return;

            savePDF(false);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            moveDate(1, true);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            moveDate(2, true);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            moveDate(3, true);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            moveDate(4, true);
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
            if (!IsBusy(true))
                RunRanking("Butikk");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedMacroRankingImport();
        }

        private void dataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MakeReport("Data");
        }

        private void lydOgBildeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MakeReport("AudioVideo");
        }

        private void teleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MakeReport("Tele");
        }

        private void butikkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MakeReport("Butikk");
        }

        private void graphPanelTop_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (!EmptyDatabase())
                {

                    DateTime d1 = topgraph.datoFra;
                    DateTime d2 = topgraph.dato;
                    int X = e.Location.X;
                    float gWidth = graphPanelTop.Width;
                    int days = (d2 - d1).Days;
                    float Hstep = gWidth / days;
                    int area = -1;
                    for (int i = 0; i < gWidth; i++)
                    {
                        if (X >= (i * Hstep) && X < (i * Hstep) + Hstep)
                            area = i;
                    }
                    
                    int a = (days - area) - 1;
                    if (area > -1 && area <= days && a > -1)
                    {
                        DateTime d = d2.AddDays(-a);
                        if (d.Month == appConfig.dbTo.Month && highlightDate != d)
                        {
                            highlightDate = d;
                            graphPanelTop.Invalidate();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void graphPanelTop_MouseLeave(object sender, EventArgs e)
        {
            highlightDate = pickerRankingDate.Value;
            graphPanelTop.Invalidate();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            loadLog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            MakeReport();
        }

        private void textBoxSearchTerm_Enter(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = button10;
        }

        private void textBoxSearchTerm_Leave(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = null;
        }

        private void dataGridTransaksjoner_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right && e.ColumnIndex == 3 && e.RowIndex != -1)
                {
                    var CellValue = dataGridTransaksjoner.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    lastRightClickValue = CellValue;
                    ContextMenuStrip Menu = new ContextMenuStrip();
                    ToolStripMenuItem MenuValue = new ToolStripMenuItem("Varegruppe: " + CellValue.ToString());
                    MenuValue.Enabled = false;
                    ToolStripMenuItem MenuFilter = new ToolStripMenuItem("Filtrer varegruppe");
                    ToolStripMenuItem MenuCopy = new ToolStripMenuItem("Kopier");
                    MenuFilter.MouseDown += new MouseEventHandler(Menu_Click);
                    MenuCopy.MouseDown += new MouseEventHandler(Menu_Click);
                    Menu.Items.AddRange(new ToolStripItem[] { MenuValue, MenuFilter, MenuCopy });
                    dataGridTransaksjoner.ContextMenuStrip = Menu;
                }
                else if (e.Button == MouseButtons.Right && e.ColumnIndex == 4 && e.RowIndex != -1)
                {
                    var CellValue = dataGridTransaksjoner.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    lastRightClickValue = CellValue;
                    ContextMenuStrip Menu = new ContextMenuStrip();
                    ToolStripMenuItem MenuValue = new ToolStripMenuItem("Varekode: " + CellValue.ToString());
                    MenuValue.Enabled = false;
                    ToolStripMenuItem MenuFilter = new ToolStripMenuItem("Søk varekode");
                    ToolStripMenuItem MenuCopy = new ToolStripMenuItem("Kopier");
                    MenuFilter.MouseDown += new MouseEventHandler(Menu_Click);
                    MenuCopy.MouseDown += new MouseEventHandler(Menu_Click);
                    Menu.Items.AddRange(new ToolStripItem[] { MenuValue, MenuFilter, MenuCopy });
                    dataGridTransaksjoner.ContextMenuStrip = Menu;
                }
                else if (e.Button == MouseButtons.Right && e.ColumnIndex == 2 && e.RowIndex != -1)
                {
                    var CellValue = dataGridTransaksjoner.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    lastRightClickValue = CellValue;
                    ContextMenuStrip Menu = new ContextMenuStrip();
                    ToolStripMenuItem MenuValue = new ToolStripMenuItem("Bilagsnr: " + CellValue.ToString());
                    MenuValue.Enabled = false;
                    ToolStripMenuItem MenuBilag = new ToolStripMenuItem("Søk bilagnr");
                    ToolStripMenuItem MenuCopy = new ToolStripMenuItem("Kopier");
                    MenuBilag.MouseDown += new MouseEventHandler(Menu_Click);
                    MenuCopy.MouseDown += new MouseEventHandler(Menu_Click);
                    Menu.Items.AddRange(new ToolStripItem[] { MenuValue, MenuBilag, MenuCopy });
                    dataGridTransaksjoner.ContextMenuStrip = Menu;
                }
                else if (e.Button == MouseButtons.Right && e.ColumnIndex == 1 && e.RowIndex != -1)
                {
                    var CellValue = dataGridTransaksjoner.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    lastRightClickValue = CellValue;
                    ContextMenuStrip Menu = new ContextMenuStrip();
                    ToolStripMenuItem MenuValue = new ToolStripMenuItem("Selger: " + CellValue.ToString());
                    MenuValue.Enabled = false;
                    ToolStripMenuItem MenuSkFilt = new ToolStripMenuItem("Filtrer selger");
                    ToolStripMenuItem MenuSk = new ToolStripMenuItem("Søk selger");
                    ToolStripMenuItem MenuCopy = new ToolStripMenuItem("Kopier");

                    var MenuAdd = new ToolStripMenuItem("Legg til..");
                    MenuAdd.DropDownItems.Add("SDA", null, this.SelgerkodeClick);
                    MenuAdd.DropDownItems.Add("AudioVideo", null, this.SelgerkodeClick);
                    MenuAdd.DropDownItems.Add("MDA", null, this.SelgerkodeClick);
                    MenuAdd.DropDownItems.Add("Tele", null, this.SelgerkodeClick);
                    MenuAdd.DropDownItems.Add("Data", null, this.SelgerkodeClick);
                    MenuAdd.DropDownItems.Add("Teknikere", null, this.SelgerkodeClick);

                    MenuSkFilt.MouseDown += new MouseEventHandler(Menu_Click);
                    MenuSk.MouseDown += new MouseEventHandler(Menu_Click);
                    MenuCopy.MouseDown += new MouseEventHandler(Menu_Click);
                    Menu.Items.AddRange(new ToolStripItem[] { MenuValue, MenuSkFilt, MenuSk, MenuCopy, MenuAdd });
                    dataGridTransaksjoner.ContextMenuStrip = Menu;
                }
                else
                    dataGridTransaksjoner.ContextMenuStrip = null;
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }

        }

        private void Menu_Click(object sender, MouseEventArgs e)
        {
            try
            {
                var cellvalue = lastRightClickValue;
                lastRightClickValue = "";
                
                var arg = sender.ToString();
                if (arg.StartsWith("Søk varekode"))
                    SearchDBold("search", false, cellvalue);
                else if (arg.StartsWith("Søk bilagnr"))
                    SearchDBold("search", false, cellvalue);
                else if (arg.StartsWith("Søk selger"))
                    SearchDBold("selgerkode", false, cellvalue);
                else if (arg.StartsWith("Filtrer varegruppe"))
                    SearchDBold("varegruppe", true, cellvalue);
                else if (arg.StartsWith("Filtrer selger"))
                    SearchDBold("selgerkode", true, cellvalue);
                else if (arg.StartsWith("Kopier"))
                    System.Windows.Forms.Clipboard.SetText(cellvalue);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            textBoxSearchTerm.Text = "";
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            LagreSelgerkoder();
        }

        private void toolStripComboBox1_DropDownClosed(object sender, EventArgs e)
        {
            OppdaterSelgerkoder();
        }

        private void listBoxSk_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                listBoxSk.SelectedIndex = listBoxSk.IndexFromPoint(e.Location);
                if (listBoxSk.SelectedIndex != -1)
                {
                    listboxContextMenu.Show();
                }
            }
        }

        private void tabControlMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            appConfig.savedTab = tabControlMain.SelectedTab.Text;
            if (tabControlMain.SelectedTab == tabControlMain.TabPages["tabPageSelgere"] && !EmptyDatabase())
            {
                InitSelgerkoder();
            }
            if (tabControlMain.SelectedTab == tabControlMain.TabPages["tabPageTrans"] && !EmptyDatabase())
            {
                InitDB();
            }
            if (tabControlMain.SelectedTab == tabControlMain.TabPages["tabPageGrafikk"] && !EmptyDatabase())
            {
                InitGraph();
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            listBoxSk.ClearSelected();
        }

        private void listBoxSk_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Loaded)
            {
                if (listBoxSk.SelectedItems.Count > 0 && comboBoxKategorier.SelectedIndex > 0)
                    button17.Enabled = true;
                else
                    button17.Enabled = false;
            }
        }

        private void comboBoxKategorier_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Loaded)
            {
                if (listBoxSk.SelectedItems.Count > 0 && comboBoxKategorier.SelectedIndex > 0)
                    button17.Enabled = true;
                else
                    button17.Enabled = false;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (listBoxSk.SelectedItems.Count > 0 && comboBoxKategorier.SelectedIndex > 0)
            {
                sender = comboBoxKategorier.SelectedItem.ToString();
                SelgerkodeClick(sender, e);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            AddSelgerkoderAuto();
        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            InitDB(false);
        }

        private void dataGridViewSk_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                if (dataGridViewSk.CurrentRow != null)
                    for (int i = 0; i < dsSk.Tables["tblSelgerkoder"].Rows.Count; i++)
                        dsSk.Tables["tblSelgerkoder"].Rows[i]["Avdeling"] = appConfig.Avdeling;
            }
            catch(DeletedRowInaccessibleException ex)
            {
                Log.n("Unntak oppstod under oppdatering av selgerliste. Exception: " + ex, Color.Red);
            }
            catch(Exception ex)
            {
                FormError errorMsg = new FormError("Unntak oppstod under oppdatering av selgerliste", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void toolStripMenuItem4_Click_1(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageSelgere;
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            ReloadDatabase();
        }

        private void sammenligningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            if (sammenligningToolStripMenuItem.Checked)
            {
                appConfig.rankingCompareLastyear = 0;
                sammenligningToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.rankingCompareLastyear = 1;
                sammenligningToolStripMenuItem.Checked = true;
            }
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void favorittAvdelingerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!appConfig.favVis)
            {
                appConfig.favVis = true;
                favorittAvdelingerToolStripMenuItem.Checked = true;
            }
            else
            {
                appConfig.favVis = false;
                favorittAvdelingerToolStripMenuItem.Checked = false;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void grafikkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!appConfig.graphVis)
            {
                appConfig.graphVis = true;
                grafikkToolStripMenuItem.Checked = true;
            }
            else
            {
                appConfig.graphVis = false;
                grafikkToolStripMenuItem.Checked = false;
            }
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";
            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            if (!IsBusy())
                velgDato();
        }

        private void leggInnEngangsMeldingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var autorankingMelding = new FormMsgAutoranking(appConfig);
            if (autorankingMelding.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SaveSettings();
                Log.n("Meldingen: " + appConfig.epostNesteMelding);
                Log.n("Melding lagret. Sendes ved nest autoranking.", Color.Green);
            }
        }

        private void buttonNotification_Click(object sender, EventArgs e)
        {
            panelNotification.Visible = false;
            labelNotificationText.Text = "";
            panelGraphNotification.Visible = false;
            labelGraphNotificationText.Text = "";

            datoPeriodeVelger = false;

            if (tabControlMain.SelectedTab == tabPageRank)
                UpdateRank(appConfig.savedPage);
            else if (tabControlMain.SelectedTab == tabPageGrafikk)
            {
                if (_graphInitialized)
                    UpdateGraph();
                else
                    InitGraph();
            }
        }

        private void button18_Click_2(object sender, EventArgs e)
        {
            if (!IsBusy())
                velgServiceCSV();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunServiceList();
        }

        private void buttonServBF_Click(object sender, EventArgs e)
        {
            moveDateService(1);
        }

        private void buttonServB_Click(object sender, EventArgs e)
        {
            moveDateService(2);
        }

        private void buttonServF_Click(object sender, EventArgs e)
        {
            moveDateService(3);
        }

        private void buttonServFF_Click(object sender, EventArgs e)
        {
            moveDateService(4);
        }

        private void dateServicePicker_DropDown(object sender, EventArgs e)
        {
            chkServicePicker = pickerServiceDato.Value;
        }

        private void dateServicePicker_CloseUp(object sender, EventArgs e)
        {
            if (chkServicePicker != pickerServiceDato.Value)
                moveDateService(0);
        }

        private void buttonServiceOppdater_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                UpdateService();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunServiceOversikt();
        }

        private void webHTML_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (graphPanelTop.Visible && !EmptyDatabase())
            {
                string kat = currentPage();

                if (kat == "Butikk" || kat == "Data" || kat == "Tele" || kat == "AudioVideo")
                    RunTopGraphUpdate(kat);
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (!_graphInitialized)
                InitGraph();
            UpdateGraph();
            _graphInitialized = true;
        }

        private void panelGrafikkPanel_Paint(object sender, PaintEventArgs e)
        {
            PaintGraph();
        }

        private void panelGrafikkPanel_Resize(object sender, EventArgs e)
        {
            PaintGraph();
        }

        private void pickerDato_Graph_CloseUp(object sender, EventArgs e)
        {
            if (chkGraphPicker != pickerDato_Graph.Value)
                moveDate(0, true);
        }

        private void pickerDato_Graph_DropDown(object sender, EventArgs e)
        {
            chkGraphPicker = pickerDato_Graph.Value;
        }

        private void buttonGraphBF_Click(object sender, EventArgs e)
        {
            moveDateGraph(1, true);
        }

        private void buttonGraphB_Click(object sender, EventArgs e)
        {
            moveDateGraph(2, true);
        }

        private void buttonGraphF_Click(object sender, EventArgs e)
        {
            moveDateGraph(3, true);
        }

        private void buttonGraphFF_Click(object sender, EventArgs e)
        {
            moveDateGraph(4, true);
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            SelectGraph("Data");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void button_GraphNettbrett_Click(object sender, EventArgs e)
        {
            SelectGraph("Nettbrett");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void button_GraphTver_Click(object sender, EventArgs e)
        {
            SelectGraph("AudioVideo");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void button_GraphMobiler_Click(object sender, EventArgs e)
        {
            SelectGraph("Tele");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (Loaded)
            {
                if (checkBox_GraphHitrateMTD.Checked)
                    appConfig.graphHitrateMTD = true;
                else
                    appConfig.graphHitrateMTD = false;
            }
        }

        private void comboBox_GraphLengde_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_GraphLengde.SelectedIndex != -1 && Loaded)
            {
                appConfig.graphLengthIndex = comboBox_GraphLengde.SelectedIndex;
                Log.d("lagret ny index.. " + appConfig.graphLengthIndex);
            }
        }

        private void listBox_GraphSelgere_Click(object sender, EventArgs e)
        {
            if (listBox_GraphSelgere.Items.Count > 0)
            {
                if (listBox_GraphSelgere.SelectedItem != null)
                    _graphSelCurrent = listBox_GraphSelgere.SelectedItem.ToString();
            }
            if (_graphSelCurrent == "ALLE")
                _graphSelCurrent = "";
        }

        private void checkBoxSKsort_Click(object sender, EventArgs e)
        {
            if (selgerkodeList.Count > 0)
            {
                if (checkBoxSKsort.Checked)
                {
                    string[] str = selgerkodeList.ToArray();
                    Array.Sort(str);
                    listBoxSk.Items.Clear();
                    listBoxSk.Items.AddRange(str);
                }
                else
                {
                    listBoxSk.Items.Clear();
                    listBoxSk.Items.AddRange(selgerkodeList.ToArray());
                }
            }

        }

        private void checkBox_GraphSort_Click(object sender, EventArgs e)
        {
            if (selgerkodeList.Count > 0)
            {
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
            }
        }

        private void button_GraphButikk_Click(object sender, EventArgs e)
        {
            SelectGraph("Butikk");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void button_GraphTjen_Click(object sender, EventArgs e)
        {
            SelectGraph("Oversikt");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void checkBox_GraphZoom_Click(object sender, EventArgs e)
        {
            if (checkBox_GraphZoom.Checked)
                appConfig.graphScreenZoom = true;
            else
                appConfig.graphScreenZoom = false;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunServiceImport();
        }

        private void buttonServiceMacro_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                delayedAutoServiceImport();
        }

        private void button8_Click_3(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void button9_Click_2(object sender, EventArgs e)
        {
            try
            {
                if (_graphInitialized && gc != null)
                {
                    using (var b = new Bitmap(appConfig.graphResX, appConfig.graphResY))
                    {
                        Bitmap graphBitmap = gc.DrawImageChunk(_graphKatCurrent, appConfig.graphResX, appConfig.graphResY, "", _graphFraDato, pickerDato_Graph.Value, true, null, false, _graphSelCurrent);
                        Clipboard.SetImage(graphBitmap);
                        Log.n("Graf lagret til utklippstavlen.");
                    }
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            TransExportCSV();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (appConfig.WindowMinimizeToTray)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    notifyIcon1.Visible = true;
                    notifyIcon1.ShowBalloonTip(3000);
                    this.ShowInTaskbar = false;
                }
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            RestoreWindow();
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            RestoreWindow();
        }

        private void RestoreWindow()
        {
            try
            {
                if (appConfig.WindowMax)
                    this.WindowState = FormWindowState.Maximized;
                else
                    this.WindowState = FormWindowState.Normal;
                this.ShowInTaskbar = true;
                notifyIcon1.Visible = false;
                notifyIcon1.BalloonTipText = "Kjører fortsatt i bakgrunnen.";
                this.TopMost = true;
                this.TopMost = false;
                this.Activate();
                Log.d("Hoved vindu satt tilbake til normal tilstand.");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void visToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RestoreWindow();
        }

        private void avsluttToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void webService_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            try
            {
                var url = e.Url.OriginalString;
                if (url.Contains("#service"))
                {
                    int index = url.IndexOf("#");
                    var s = url.Substring(index + 8, url.Length - index - 8);
                    int serviceID = 0;
                    if (s.Length > 0)
                        serviceID = Convert.ToInt32(s);
                    else
                    {
                        Log.n("Ingen service å søke opp.", Color.Red);
                        return;
                    }
                    MakeServiceDetails(serviceID);
                }
                if (url.Contains("#list_"))
                {
                    int index = url.IndexOf("#");
                    var status = url.Substring(index + 6, url.Length - index - 6);

                    status = status.Replace("%20", " ");
                    if (status == "alle")
                        status = "";
                    MakeServiceList(status);
                }

            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil ved service navigering", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void dataGridViewSk_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(e.FormattedValue.ToString()) && e.ColumnIndex < 3)
                {
                    Log.Status("Feltet kan ikke være tomt! Eller har du ikke oppdatert databasen?", Color.Red);
                    e.Cancel = true;
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                dataGridViewSk[e.ColumnIndex, e.RowIndex].Value = "";
            }
        }

        private void dataGridViewSk_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Log.Status("");
        }

        private void serviceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageService;
        }

        private void nullstillBehandletMarkeringToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                service.NullstillBehandlet();
                UpdateService();
            }
        }

        private void rabattToolStripMenuItem_Click(object sender, EventArgs e)
        {
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";

            if (appConfig.kolRabatt)
            {
                appConfig.kolRabatt = false;
                rabattToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.kolRabatt = true;
                rabattToolStripMenuItem.Checked = true;
            }

            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void epostToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormEmailAddressbook emailForm = new FormEmailAddressbook();
            emailForm.ShowDialog(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormKrav form = new FormKrav(this);
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SaveSettings();
            }
            OppdaterSelgerkoder();
            ClearHash("Oversikt");
        }

        private void startAutoLagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedAutoStore();
        }

        private void sistMånedSammenligningToolStripMenuItem_Click(object sender, EventArgs e)
        {
            processing.SetVisible = true;
            processing.SetText = "Oppdaterer..";

            if (sistMånedSammenligningToolStripMenuItem.Checked)
            {
                appConfig.rankingCompareLastmonth = 0;
                sistMånedSammenligningToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.rankingCompareLastmonth = 1;
                sistMånedSammenligningToolStripMenuItem.Checked = true;
            }

            SaveSettings();
            Reload(true);
            processing.SetVisible = false;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("Obsolete");
        }

        private void buttonOppdaterLager_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                UpdateStore();
        }

        private void pickerLagerDato_DropDown(object sender, EventArgs e)
        {
            chkStorePicker = pickerLagerDato.Value;
        }

        private void pickerLagerDato_CloseUp(object sender, EventArgs e)
        {
            if (chkStorePicker != pickerLagerDato.Value)
                moveStoreDate(0, true);
        }

        private void buttonLagerUkuListe_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("ObsoleteList");
        }

        private void button40_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                ImportWobsoleteCsvZip();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedAutoStore();
        }

        private void buttonLagerBF_Click(object sender, EventArgs e)
        {
            moveStoreDate(1, true);
        }

        private void buttonLagerB_Click(object sender, EventArgs e)
        {
            moveStoreDate(2, true);
        }

        private void buttonLagerF_Click(object sender, EventArgs e)
        {
            moveStoreDate(3, true);
        }

        private void buttonLagerFF_Click(object sender, EventArgs e)
        {
            moveStoreDate(4, true);
        }

        private void button18_Click_4(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("ObsoleteImports");
        }

        private void button18_Click_5(object sender, EventArgs e)
        {
            if (!IsBusy())
                velgLagerViewpointDato();
        }

        private void button43_Click(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void button42_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                savePDF(false);
        }

        private void lagreSideSomPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                savePDF(false);
        }

        private void rankingirankcsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunImport();
        }

        private void serviceToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunServiceImport();
        }

        private void lagerwobsoletecsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportWobsoleteCsvZip();
        }

        private void velgIrankcsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                velgCSV();
        }

        private void velgIservcsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                velgServiceCSV();
        }

        private void ShowHideOverlayBottom(object sender, EventArgs e)
        {
            if (panelOverlayBottom.Visible)
                panelOverlayBottom.Visible = false;
            else
                panelOverlayBottom.Visible = true;
        }

        private void åpneSomPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void velgWobsoletezipcsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            velgObsoleteCSV();
        }

        void LogCopyAction(object sender, EventArgs e)
        {
            if (richLog.SelectedText.Length > 0)
                Clipboard.SetText(richLog.SelectedText);
        }

        void LogSelectAllAction(object sender, EventArgs e)
        {
            richLog.SelectAll();
        }

        private void panelGrafikkPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (_graphInitialized && gc != null)
                {
                    ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();
                    MenuItem menuItem = new MenuItem("Kopier");
                    menuItem.Click += new EventHandler(GraphCopyAction);
                    contextMenu.MenuItems.Add(menuItem);
                    panelGrafikkPanel.ContextMenu = contextMenu;
                }
            }
        }

        void GraphCopyAction(object sender, EventArgs e)
        {
            try
            {
                if (_graphInitialized && gc != null)
                {
                    using (var b = new Bitmap(appConfig.graphResX, appConfig.graphResY))
                    {
                        Bitmap graphBitmap = gc.DrawImageChunk(_graphKatCurrent, appConfig.graphResX, appConfig.graphResY, "", _graphFraDato, pickerDato_Graph.Value, true, null, false, _graphSelCurrent);
                        Clipboard.SetImage(graphBitmap);
                        Log.n("Graf lagret til utklippstavlen.");
                    }
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                RunRanking("KnowHow");
        }

        private void button26_Click(object sender, EventArgs e)
        {
            SelectGraph("KnowHow");
            if (_graphInitialized && !EmptyDatabase())
                UpdateGraph();
        }

        private void rankingKnowHowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
            if (!IsBusy(true))
                RunRanking("KnowHow");
        }

        private void buttonLister_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                RunRanking("Lister");
        }

        private void RunCreateHtml()
        {
            if (!bwCreateHtml.IsBusy)
                bwCreateHtml.RunWorkerAsync();
        }

        private void richLog_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            string link = e.LinkText.Replace((char)160, ' ');
            try
            {
                Process.Start(link);
            }
            catch(FileNotFoundException)
            {
                Log.Alert("Dokumentet " + link + " finnes ikke.\nFilen kan være slettet, låst eller ødelagt");
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("Klarte ikke åpne link. Feilmelding: " + ex.Message);
            }
        }

        private void nyBudsjettToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                OpenBudgetSettings();
        }

        private void buttonBudgetUpdate_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                UpdateBudget();
        }

        private void buttonBudgetMda_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.MDA);
        }

        private void buttonBudgetAv_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.AudioVideo);
        }

        private void buttonBudgetSda_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.SDA);
        }

        private void buttonBudgetTele_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Tele);
        }

        private void buttonBudgetData_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Data);
        }

        private void buttonBudgetKasse_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Cross);
        }

        private void pickerBudget_CloseUp(object sender, EventArgs e)
        {
            if (chkBudgetPicker != pickerBudget.Value)
                moveBudgetDate(0, true);
        }

        private void pickerBudget_DropDown(object sender, EventArgs e)
        {
            chkBudgetPicker = pickerBudget.Value;
        }

        private void webBudget_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            NavigateWeb(e);
        }

        private void vinnproduktToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var formVinn = new FormVinnprodukt(this);

            if (formVinn.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                ClearHash("Vinnprodukter");
                vinnprodukt = new Vinnprodukt(this);
            }
        }

        private void buttonBudgetKasse_Click_1(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Kasse);
        }

        private void buttonBudgetAftersales_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Aftersales);
        }

        private void buttonRankVinnprodukter_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Vinnprodukter");
        }

        private void buttonBudgetMdaSda_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.MDASDA);
        }

        private void buttonBudgetFF_Click(object sender, EventArgs e)
        {
            moveBudgetDate(4, true);
        }

        private void buttonBudgetBF_Click(object sender, EventArgs e)
        {
            moveBudgetDate(1, true);
        }

        private void rankingToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageRank;
        }

        private void buttonBudgetF_Click(object sender, EventArgs e)
        {
            moveBudgetDate(3, true);
        }

        private void buttonBudgetB_Click(object sender, EventArgs e)
        {
            moveBudgetDate(2, true);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                OpenBudgetSettings();
        }

        private void lagreBudsjettPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                saveBudgetPDF(true);
        }

        private void lagerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageStore;
        }

        private void budsjettToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            tabControlMain.SelectedTab = tabPageBudget;
        }

        private void buttonBudgetButikk_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Butikk);
        }

        private void varekoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new FormVarekoder(this);
            if (form.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                SaveSettings();

            form.Dispose();
        }

        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            if (openXml != null)
                openXml.OpenDocument(pickerRankingDate.Value);
        }

        private void buttonLagerExcel_Click(object sender, EventArgs e)
        {
            if (openXml != null)
                openXml.OpenDocument(pickerLagerDato.Value);
        }

        private void buttonLagerWeekly_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("LagerUkeAnnonser");
        }

        private void buttonLagerPrisguide_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("LagerPrisguide");
        }

        private void importerPrisguidenoProdukterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                OnlineImporter online = new OnlineImporter(this);
                worker = new BackgroundWorker();
                online.StartAsyncPrisguideImport(worker);
            }
        }

        private void importerUkenyttToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                OnlineImporter online = new OnlineImporter(this);
                worker = new BackgroundWorker();
                online.StartAsyncWeeklyImport(worker);
            }
        }

        private void webLager_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            NavigateStoreWeb(e);
        }

        private void buttonLagerWeeklyOverview_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("LagerUkeAnnonserOversikt");
        }

        private void buttonLagerPrisguideOverview_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunStore("LagerPrisguideOversikt");
        }

        private void buttonLagerOppUkeannonser_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                OnlineImporter online = new OnlineImporter(this);
                worker = new BackgroundWorker();
                online.StartAsyncWeeklyImport(worker);
            }
        }

        private void buttonLagerOppPrisguide_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                OnlineImporter online = new OnlineImporter(this);
                worker = new BackgroundWorker();
                online.StartAsyncPrisguideImport(worker);
            }
        }

        private void hentDagensBudsjettToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                BudgetImporter kpi = new BudgetImporter(this, DateTime.Now);
                worker = new BackgroundWorker();
                kpi.StartAsyncDownloadBudget(worker, false);
            }
        }

        private void inkluderBudsjettMålIKveldstallToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (appConfig.dailyBudgetIncludeInQuickRanking)
                appConfig.dailyBudgetIncludeInQuickRanking = false;
            else
                appConfig.dailyBudgetIncludeInQuickRanking = true;

            inkluderBudsjettMålIKveldstallToolStripMenuItem.Checked = appConfig.dailyBudgetIncludeInQuickRanking;

            SaveSettings();
        }

        private void oppdaterAutomatiskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (appConfig.dailyBudgetQuickRankingAutoUpdate)
                appConfig.dailyBudgetQuickRankingAutoUpdate = false;
            else
                appConfig.dailyBudgetQuickRankingAutoUpdate = true;

            oppdaterAutomatiskToolStripMenuItem.Checked = appConfig.dailyBudgetQuickRankingAutoUpdate;

            SaveSettings();
        }

        private void buttonAvdMacroImport_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedMacroRankingImport();
        }

        private void åpneIExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true) && openXml != null)
            {
                string currentTab = readCurrentTab();
                if (currentTab.Equals("Avdelinger"))
                    openXml.OpenDocument(pickerRankingDate.Value);
                else if (currentTab.Equals("Store"))
                    openXml.OpenDocument(pickerLagerDato.Value);
                else if (currentTab.Equals("Ranking"))
                    openXml.OpenDocument(pickerRankingDate.Value);
                else
                    Log.n("Kan ikke åpne regneark for gjeldene side. Velg en annen.", Color.Red);
            }
        }

        private void buttonBudgetMacro_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                startDelayedDailyImport();
        }

        private void buttonBudgetDaily_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.Daglig);
        }

        private void buttonAvdTjenester_Click_1(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Tjenester");
        }

        private void buttonAvdSnittpriser_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunRanking("Snittpriser");
        }

        private void buttonBudgetActionImportCsv_Click(object sender, EventArgs e)
        {
            RunImport();
        }

        private void buttonBudgetActionOpenExcel_Click(object sender, EventArgs e)
        {
            if (openXml != null)
                openXml.OpenDocument(pickerBudget.Value);
        }

        private void buttonBudgetActionMacroImport_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedMacroRankingImport();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void buttonBudgetActionOpenPdf_Click(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void buttonStoreOpenPdf_Click(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void button4_Click_3(object sender, EventArgs e)
        {
            RunOpenPDF();
        }

        private void kjørAutorankingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                delayedAutoRanking();
        }

        private void kjørKveldstankingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                startDelayedAutoQuickImport();
        }

        private void hentServicerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedAutoServiceImport();
        }

        private void hentLagervarerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedAutoStore();
        }

        private void hentTransaksjonerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy(true))
                delayedMacroRankingImport();
        }

        private void makroInnstillingerToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                FormSettingsMacro form = new FormSettingsMacro(this);
                if (form.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    SaveSettings();
                    UpdateUi();
                }

                form.Dispose();
            }
        }

        private void budsjettToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                OpenBudgetSettings();
        }

        private void buttonBudgetActionBudgetImport_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                BudgetImporter importer = new BudgetImporter(this, DateTime.Now);
                worker = new BackgroundWorker();
                importer.StartAsyncDownloadBudget(worker, false);
            }
        }

        private void buttonBudgetAllSalesRep_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
                RunBudget(BudgetCategory.AlleSelgere);
        }

        private void oppdaterEANDatabasenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                if (Log.Alert("Informasjon:\n\nEksporter EAN koder fra meny 200 i Elguide. Velg kategori 1 til 9-99-99.\n\nKlikk OK for å fortsette.",
                    "EAN importering", MessageBoxButtons.OKCancel, MessageBoxIcon.None) == DialogResult.OK)
                {
                    AppManager app = new AppManager(this);
                    app.ImportEan();
                }
            }
        }

        private void slettAlleEANKoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                database.tableEan.Reset();
                Log.n("EAN databasen er nullstilt", Color.Green);
            }
        }

        private void påAutoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (påAutoToolStripMenuItem.Checked)
            {
                appConfig.blueServerIsEnabled = false;
                avToolStripMenuItem.Checked = true;
                påAutoToolStripMenuItem.Checked = false;
            }
            else
            {
                appConfig.blueServerIsEnabled = true;
                avToolStripMenuItem.Checked = false;
                påAutoToolStripMenuItem.Checked = true;
            }

            SaveSettings();
            ReloadStore();
            UpdateUi();
        }

        private void avToolStripMenuItem_Click(object sender, EventArgs e)
        {
            appConfig.blueServerIsEnabled = avToolStripMenuItem.Checked;
            if (appConfig.blueServerIsEnabled)
            {
                avToolStripMenuItem.Checked = false;
                påAutoToolStripMenuItem.Checked = true;
            }
            else
            {
                avToolStripMenuItem.Checked = true;
                påAutoToolStripMenuItem.Checked = false;
            }

            SaveSettings();
            ReloadStore();
            UpdateUi();
        }

        private void oppdaterAppDatabaserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy())
            {
                AppManager appMng = new AppManager(this);
                appMng.UpdateAllAsync();
            }
        }

        private void checkBoxLogDebugSql_CheckedChanged(object sender, EventArgs e)
        {
            appConfig.debugSql = checkBoxLogDebugSql.Checked;
        }

        private void checkBoxLogDebug_CheckedChanged(object sender, EventArgs e)
        {
            appConfig.debug = checkBoxLogDebug.Checked;
        }

        private void oversiktSelgereToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsBusy() && openXml != null)
            {
                openXml.CreateAndOpenXml(OpenXML.DOC_ALL_SALES_REP);
            }
        }
    }
}
