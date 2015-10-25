using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace KGSA
{
    [Serializable]
    public class AppSettings
    {
        #region Online Importer settings
        /// Enable / Disable automatic online importing of weekly products and popular prisguide products
        public bool onlineImporterAuto { get; set; }
        /// Number of pages to import at once. Every page contains 30 produkts: 30 x n = Total number of products to import
        public int onlinePrisguidePagesToImport { get; set; }
        #endregion

        #region Daily Budget settings

        /// Enable / Disable inclusion of daily budget in quick ranking
        public bool dailyBudgetIncludeInQuickRanking { get; set; }
        /// Enable / Disable automatic updates of todays budget
        public bool dailyBudgetQuickRankingAutoUpdate { get; set; }

        /// Maximum number of headers to search in inbox
        public int dailyBudgetMaxHeaders { get; set; }

        /// Maximum download attempts of emails
        public int dailyBudgetMaxAttempts { get; set; }

        #endregion

        #region Bluetooth settings

        /// Decides if bluetooth server is automatically started and listening for clients
        public bool blueServerIsEnabled { get; set; }
        /// Date when the App database was updated last
        public DateTime blueServerDatabaseUpdated { get; set; }

        /// Will not accept MScanner apps of earlier versions then this to connect to the bluetooth server
        public double blueServerMinimumAcceptedVersion { get; set; }

        /// Date for the last item in the inventory database
        public DateTime blueInventoryLastDate { get; set; }

        /// Date for when the data was exported from Elguide
        public DateTime blueInventoryExportDate { get; set; }

        /// Date for the last item in the inventory database
        public DateTime blueProductLastDate { get; set; }

        /// Date for when the data was exported from Elguide
        public DateTime blueProductExportDate { get; set; }

        /// Date for when the EAN database was last updated.
        public DateTime blueEanUpdateDate { get; set; }
        #endregion
        public DateTime dbFrom { get; set; }
        public DateTime dbTo { get; set; }
        public DateTime dbObsoleteUpdated { get; set; }
        public DateTime dbStoreFrom { get; set; }
        public DateTime dbStoreTo { get; set; }
        public DateTime dbStoreViewpoint { get; set; }
        public DateTime vinnFrom { get; set; }
        public DateTime vinnTo { get; set; }
        public int rankingAvdelingMode { get; set; } // 0 = Dag, 1 = MTD, 2 = Bonus, 3 = YTD, 4 = Uke, 5 = År
        public bool rankingAvdelingShowAll { get; set; }
        public List<string> avdelingerListAlle { get; set; }
        public bool experimental { get; set; }
        public bool vinnDatoFraTil { get; set; }
        public bool vinnSisteDagVare { get; set; }
        public bool vinnVisVarekoder { get; set; }
        public bool vinnVisVarekoderExtra { get; set; }
        public bool vinnEnkelModus { get; set; }
        public bool vinnVisAftersales { get; set; }
        public bool showAdvancedFunctions { get; set; }
        public decimal budgetChartPostWidth { get; set; }
        public int budgetChartMinPosts { get; set; } // Minimum antall stolper i graf
        public bool budgetChartShowEfficiency { get; set; }
        public bool budgetChartShowQuality { get; set; }
        public bool budgetIsolateCrossSalesRep { get; set; }
        public bool budgetInclAllSalesRepUnderCross { get; set; }
        public bool budgetShowMda { get; set; }
        public bool budgetShowAudioVideo { get; set; }
        public bool budgetShowSda { get; set; }
        public bool budgetShowTele { get; set; }
        public bool budgetShowData { get; set; }
        public bool budgetShowCross { get; set; }
        public bool budgetShowKasse { get; set; }
        public bool budgetShowAftersales { get; set; }
        public bool budgetShowMdasda { get; set; }
        public bool budgetShowButikk { get; set; }
        public bool useSqlCache { get; set; }
        public bool macroShowWarning { get; set; }
        public string importSetting { get; set; }
        public bool storeCompareMtd { get; set; }
        public bool storeShowStoreTwo { get; set; }
        public int storeMaxAgePrizes { get; set; }
        public int storeObsoleteFilterPercent { get; set; }
        public bool storeObsoleteFilterMainStoreOnly { get; set; }
        public string storeObsoleteSortBy { get; set; }
        public bool storeObsoleteSortAsc { get; set; }
        public int storeObsoleteFilterMax { get; set; }
        public bool storeObsoleteListMainProductsOnly { get; set; }
        public int rankingCompareLastyear { get; set; } // 0 = Ikke vis fjoråret, 1 = Vis fjoråret, 2 = vis fjoråret detaljert
        public int rankingCompareLastmonth { get; set; } // 0 = Ikke vis forrige måned, 1 = Vis forrige måned, 2 = vis forrige måned detaljert
        public bool rankingShowLastWeek { get; set; } // Vis hele sist uke
        public bool epostIncService { get; set; } // Inkluder service i automatisk ranking utsending
        public int WindowLocationX { get; set; }
        public int WindowLocationY { get; set; }
        public int WindowSizeX { get; set; }
        public int WindowSizeY { get; set; }
        public bool WindowMax { get; set; }
        public bool WindowMinimizeToTray { get; set; }
        public bool WindowExitToTray { get; set; }
        public bool debug { get; set; }
        public bool debugSql { get; set; }
        public bool showTrivia { get; set; }
        public bool pdfExpandedGraphs { get; set; }
        public bool pdfVisButikk { get; set; }
        public bool pdfVisKnowHow { get; set; }
        public bool pdfVisData { get; set; }
        public bool pdfVisAudioVideo { get; set; }
        public bool pdfVisTele { get; set; }
        public bool pdfVisLister { get; set; }
        public bool pdfVisVinnprodukter { get; set; }
        public bool pdfVisTjenester { get; set; }
        public bool pdfVisSnittpriser { get; set; }
        public bool pdfVisBudsjett { get; set; }
        public bool pdfVisService { get; set; }
        public bool pdfVisObsoleteList { get; set; }
        public bool pdfVisObsolete { get; set; }
        public bool pdfVisWeekly { get; set; }
        public bool pdfVisPrisguide { get; set; }
        public bool pdfVisOversikt { get; set; }
        public bool pdfVisToppselgere { get; set; }
        public bool pdfLandscape { get; set; }
        public bool pdfTableOfContent { get; set; }
        public string csvElguideExportFolder { get; set; }
        public bool pdfExport { get; set; }
        public string pdfExportFolder { get; set; }
        public bool csvMake { get; set; }
        public bool macroImportQuickSales { get; set; }
        public bool chainElkjop { get; set; }
        public string strButikk { get; set; }
        public string strKnowHow { get; set; }
        public string strData { get; set; }
        public string strAudioVideo { get; set; }
        public string strTele { get; set; }
        public string strOversikt { get; set; }
        public string strTjenester { get; set; }
        public string strSnittpriser { get; set; }
        public string strBudgetMda { get; set; }
        public string strBudgetAv { get; set; }
        public string strBudgetSda { get; set; }
        public string strBudgetTele { get; set; }
        public string strBudgetData { get; set; }
        public string strBudgetCross { get; set; }
        public string strBudgetKasse { get; set; }
        public string strBudgetAftersales { get; set; }
        public string strBudgetMdasda { get; set; }
        public string strBudgetButikk { get; set; }
        public string strBudgetDaily { get; set; }
        public string strBudgetAllSales { get; set; }
        public string strToppselgere { get; set; }
        public string strLister { get; set; }
        public string strVinnprodukter { get; set; }
        public string strObsolete { get; set; }
        public string strObsoleteList { get; set; }
        public string strObsoleteImports { get; set; }
        public string strLagerWeekly { get; set; }
        public string strLagerWeeklyOverview { get; set; }
        public string strLagerPrisguide { get; set; }
        public string strLagerPrisguideOverview { get; set; }
        public string strServiceOversikt { get; set; }
        public string strServiceList { get; set; }
        public string savedPage { get; set; }
        public string savedAvdPage { get; set; }
        public BudgetCategory savedBudgetPage { get; set; }
        public string savedStorePage { get; set; }
        public string savedTab { get; set; }
        public int Avdeling { get; set; }
        public bool kolForkort { get; set; }
        public int kolForkortLengde { get; set; }
        [XmlArrayItem("Varekode")]
        public List<VarekodeList> varekoder { get; set; }
        #region E-mail settings
        public string epostNesteMelding { get; set; }
        public string epostAvsender { get; set; }
        public string epostAvsenderNavn { get; set; }
        public string epostBrukernavn { get; set; }
        public string epostMisc { get; set; }
        public string epostPassHash { get; set; }
        public int epostSMTPport { get; set; }
        public bool epostSMTPssl { get; set; }
        public string epostSMTPserver { get; set; }
        public string epostPOP3server { get; set; }
        public string epostPOP3username { get; set; }
        public string epostPOP3password { get; set; }
        public int epostPOP3port { get; set; }
        public bool epostPOP3ssl { get; set; }
        public bool epostBrukBcc { get; set; }
        public string epostEmne { get; set; }
        public string epostEmneQuick { get; set; }
        public int epostHour { get; set; }
        public int epostMinute { get; set; }
        public int epostHourQuick { get; set; }
        public int epostMinuteQuick { get; set; }
        public int epostHourQuickSaturday { get; set; }
        public int epostMinuteQuickSaturday { get; set; }
        public int epostHourQuickSunday { get; set; }
        public int epostMinuteQuickSunday { get; set; }
        public string epostBody { get; set; }
        public string epostBodyQuick { get; set; }
        public bool epostTimerQuickSaturday { get; set; }
        public bool epostTimerQuickSunday { get; set; }
        #endregion
        public bool autoRank { get; set; }
        public bool autoQuick { get; set; }
        public int AutoStoreHour { get; set; }
        public int AutoStoreMinute { get; set; }
        public bool unlock { get; set; }
        public string unlockArg { get; set; }
        public string favAvdeling { get; set; }
        public bool epostOnlySendUpdated { get; set; } // Hvis sann, bare send hvis vi har gårsdagens tall

        // Ny graf innstillinger start
        public float graphScreenDPI { get; set; }
        public bool graphHitrateMTD { get; set; }
        public int graphLengthIndex { get; set; }
        public bool graphScreenZoom { get; set; }
        // Ny graf innstillinger slutt

        // Ny Service auto import start
        public int serviceAutoImportMinutter { get; set; }
        public int serviceAutoImportFraIndex { get; set; }
        public int serviceAutoImportTilIndex { get; set; }
        public string serviceEgenServiceFilter { get; set; }
        public bool serviceShowTrend { get; set; }
        public bool serviceShowHistory { get; set; }
        public bool serviceShowHistoryGraph { get; set; }
        public int serviceHistoryDays { get; set; }
        public bool serviceFerdigServiceStats { get; set; }
        public int serviceFerdigServiceStatsAntall { get; set; }
        // Ny service auto import slutt
        public int graphLimit { get; set; }
        public int graphResX { get; set; }
        public int graphResY { get; set; }
        public int graphWidth { get; set; }
        public int graphDager { get; set; }
        public bool favVis { get; set; }
        public bool graphVis { get; set; }
        public bool graphButikk { get; set; }
        public bool graphKnowHow { get; set; }
        public bool graphData { get; set; }
        public bool graphAudioVideo { get; set; }
        public bool graphTele { get; set; }
        public bool graphOversikt { get; set; }
        public bool graphAdvanced { get; set; }
        public bool graphExtra { get; set; }
        public bool histogramVis { get; set; }
        public bool kolKravData { get; set; }
        public bool kolKravAudioVideo { get; set; }
        public bool kolKravTele { get; set; }
        public bool kolKravNettbrett { get; set; }
        public decimal kravHitrateData { get; set; }
        public decimal kravHitrateAudioVideo { get; set; }
        public decimal kravHitrateTele { get; set; }
        public decimal kravHitrateNettbrett { get; set; }
        public bool kolProv { get; set; }
        public bool kolInntjen { get; set; }
        public bool kolSalgspris { get; set; }
        public bool kolRabatt { get; set; }
        public bool kolVarekoder { get; set; }
        public int sortIndex { get; set; }
        public bool combineFriSat { get; set; }
        public bool ignoreSunday { get; set; }
        public string visningNull { get; set; }
        public bool visningJevnfarge { get; set; }
        public string macroElguide { get; set; }
        public decimal macroLatency { get; set; }
        public bool AutoService { get; set; }
        public bool AutoStore { get; set; }
        /// <summary>
        /// Topp selgere sammenligning innstilling:
        /// 0 = Beste selger inntjent/omsatt fjoråret (MTD)
        /// 1 = Beste selger antall fjoråret (MTD)
        /// 2 = Beste selger inntjent/omsatt fjoråret (Totalt)
        /// 3 = Beste selger antall fjoråret (Totalt)
        /// </summary>
        public int bestofCompareIndex { get; set; }
        public bool bestofCompareChange { get; set; }
        public bool listerVisInntjen { get; set; }
        public bool listerVisAccessories { get; set; }
        public int[] accessoriesListMda { get; set; }
        public int[] accessoriesListAv { get; set; }
        public int[] accessoriesListSda { get; set; }
        public int[] accessoriesListTele { get; set; }
        public int[] accessoriesListData { get; set; }
        public int[] mainproductListMda { get; set; }
        public int[] mainproductListAv { get; set; }
        public int[] mainproductListSda { get; set; }
        public int[] mainproductListTele { get; set; }
        public int[] mainproductListData { get; set; }
        public int listerMaxLinjer { get; set; }
        public int bestofTallTA { get; set; }
        public int bestofTallFinans { get; set; }
        public int bestofTallStrom { get; set; }
        public int bestofTallTjen { get; set; }
        public int bestofTallInntjen { get; set; }
        public bool bestofSortInntjenSecondary { get; set; }
        public bool bestofSortTjenesterSecondary { get; set; }
        public bool bestofVisBesteLastYear { get; set; }
        public bool bestofVisBesteLastYearTotal { get; set; }
        public bool bestofVisBesteLastOpenday { get; set; }
        public bool bestofHoppoverKasse { get; set; }
        public decimal pdfZoom { get; set; }
        public bool oversiktHideAftersales { get; set; }
        public bool oversiktHideKitchen { get; set; }
        public bool oversiktKravVis { get; set; }
        public bool oversiktKravFinans { get; set; }
        public bool oversiktKravMod { get; set; }
        public bool oversiktKravStrom { get; set; }
        public bool oversiktKravRtgsa { get; set; }
        public bool oversiktKravFinansAntall { get; set; }
        public bool oversiktKravModAntall { get; set; }
        public bool oversiktKravStromAntall { get; set; }
        public bool oversiktKravRtgsaAntall { get; set; }
        public bool oversiktFilterToDepartments { get; set; }
        public bool oversiktKravMtd { get; set; }
        public bool oversiktKravMtdShowTarget { get; set; }

        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color1 { get; set; }
        public bool color1inv { get; set; }
        public decimal color1max { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color2 { get; set; }
        public bool color2inv { get; set; }
        public decimal color2max { get; set; }
        public decimal color2min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color3 { get; set; }
        public bool color3inv { get; set; }
        public decimal color3max { get; set; }
        public decimal color3min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color4 { get; set; }
        public bool color4inv { get; set; }
        public decimal color4max { get; set; }
        public decimal color4min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color5 { get; set; }
        public bool color5inv { get; set; }
        public decimal color5max { get; set; }
        public decimal color5min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color color6 { get; set; }
        public bool color6inv { get; set; }
        public decimal color6min { get; set; }

        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor1 { get; set; }
        public bool ncolor1inv { get; set; }
        public decimal ncolor1max { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor2 { get; set; }
        public bool ncolor2inv { get; set; }
        public decimal ncolor2max { get; set; }
        public decimal ncolor2min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor3 { get; set; }
        public bool ncolor3inv { get; set; }
        public decimal ncolor3max { get; set; }
        public decimal ncolor3min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor4 { get; set; }
        public bool ncolor4inv { get; set; }
        public decimal ncolor4max { get; set; }
        public decimal ncolor4min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor5 { get; set; }
        public bool ncolor5inv { get; set; }
        public decimal ncolor5max { get; set; }
        public decimal ncolor5min { get; set; }
        [XmlElement(Type = typeof(XmlColor))]
        public System.Drawing.Color ncolor6 { get; set; }
        public bool ncolor6inv { get; set; }
        public decimal ncolor6min { get; set; }

        public AppSettings()
        {
            // DEFAULT SETTINGS

            // online settings defaults
            onlineImporterAuto = false;
            onlinePrisguidePagesToImport = 1;

            // bluetooth server defaults
            blueServerIsEnabled = false;
            blueServerDatabaseUpdated = FormMain.rangeMin;
            blueServerMinimumAcceptedVersion = 1;
            blueInventoryLastDate = FormMain.rangeMin;
            blueProductLastDate = FormMain.rangeMin;
            blueEanUpdateDate = FormMain.rangeMin;

            if (WindowSizeX < 994) { WindowSizeX = 994; }
            if (WindowSizeY < 736) { WindowSizeY = 736; }
            if (WindowLocationX <= 0) { WindowLocationX = 50; }
            if (WindowLocationY <= 0) { WindowLocationY = 50; }
            WindowMinimizeToTray = true;
            importSetting = "Full";
            rankingCompareLastyear = 0;
            serviceFerdigServiceStatsAntall = 5;
            storeMaxAgePrizes = 50;
            storeObsoleteListMainProductsOnly = true;
            storeShowStoreTwo = true;
            savedPage = "Data";
            savedBudgetPage = BudgetCategory.MDA;
            savedStorePage = "Obsolete";
            savedTab = "Ranking ";
            listerMaxLinjer = 25;
            useSqlCache = false;
            rankingAvdelingMode = 1;
            rankingAvdelingShowAll = true;
            avdelingerListAlle = new List<string>() { };
            varekoder = new List<VarekodeList>();

            // Budget
            budgetChartPostWidth = 70;
            budgetChartMinPosts = 8;
            budgetChartShowEfficiency = true;
            budgetChartShowQuality = true;
            budgetShowMda = true;
            budgetShowAudioVideo = true;
            budgetShowSda = false;
            budgetShowTele = true;
            budgetShowData = true;
            budgetShowCross = true;
            budgetShowKasse = false;
            budgetShowAftersales = false;
            budgetShowMdasda = false;

            vinnFrom = DateTime.Now;
            vinnTo = DateTime.Now;

            mainproductListMda = new int[8] { 131, 132, 134, 135, 136, 143, 144, 145 };
            mainproductListAv = new int[2] { 224, 273 };
            mainproductListSda = new int[2] { 301, 346 };
            mainproductListTele = new int[2] { 431, 447 };
            mainproductListData = new int[3] { 531, 533, 534 };

            accessoriesListMda = new int[1] { 195 };
            accessoriesListAv = new int[1] { 214 };
            accessoriesListSda = new int[1] { 395 };
            accessoriesListTele = new int[1] { 487 };
            accessoriesListData = new int[3] { 552, 569, 589 };

            dbFrom = FormMain.rangeMin;
            dbTo = FormMain.rangeMin;
            dbStoreFrom = FormMain.rangeMin;
            dbStoreTo = FormMain.rangeMin;
            dbObsoleteUpdated = FormMain.rangeMin;
            chainElkjop = true;
            bestofSortInntjenSecondary = false;
            bestofSortTjenesterSecondary = false;
            bestofVisBesteLastYear = false;
            bestofVisBesteLastYearTotal = false;
            bestofCompareChange = true;
            AutoStoreHour = 8;
            AutoStoreMinute = 0;
            storeMaxAgePrizes = 30;
            storeObsoleteSortBy = "tblUkurans.UkuransProsent";
            storeObsoleteSortAsc = false;
            storeObsoleteFilterMax = 30;
            epostPOP3password = "094109029165239198160022101139180078146211183228"; // default passord: user1

            color1 = Color.Maroon;
            color1inv = true;
            color1max = 16.999M;
            color2 = Color.Red;
            color2inv = true;
            color2max = 24.999M;
            color2min = 17M;
            color3 = Color.Orange;
            color3inv = true;
            color3max = 29.999M;
            color3min = 25M;
            color4 = Color.Yellow;
            color4inv = false;
            color4max = 39.999M;
            color4min = 30M;
            color5 = Color.Green;
            color5inv = true;
            color5max = 99.999M;
            color5min = 40M;
            color6 = Color.Green;
            color6inv = false;
            color6min = 100M;

            ncolor1 = Color.Maroon;
            ncolor1inv = true;
            ncolor1max = 4.999M;
            ncolor2 = Color.Red;
            ncolor2inv = true;
            ncolor2max = 6.999M;
            ncolor2min = 5M;
            ncolor3 = Color.Orange;
            ncolor3inv = true;
            ncolor3max = 9.999M;
            ncolor3min = 7M;
            ncolor4 = Color.Yellow;
            ncolor4inv = false;
            ncolor4max = 14.999M;
            ncolor4min = 10M;
            ncolor5 = Color.Green;
            ncolor5inv = true;
            ncolor5max = 99.999M;
            ncolor5min = 15M;
            ncolor6 = Color.Green;
            ncolor6inv = false;
            ncolor6min = 100M;

            macroLatency = 1;

            pdfVisToppselgere = true;
            pdfVisOversikt = true;
            pdfVisButikk = true;
            pdfVisData = true;
            pdfVisAudioVideo = true;
            pdfVisTele = true;
            
            pdfExportFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            bestofTallTA = 10;
            bestofTallStrom = 10;
            bestofTallFinans = 10;
            bestofTallTjen = 10;
            bestofTallInntjen = 10;

            epostNesteMelding = "";
            unlockArg = "-p user01";
            epostSMTPport = 25;
            epostSMTPserver = "smtp.elkjop.no";
            epostBody = "Se vedlegg..";
            epostBodyQuick = "Se vedlegg..";
            epostAvsender = "kgsa@elkjop.no";
            kolProv = true;
            kolInntjen = true;
            csvElguideExportFolder = @"C:\";
            epostPOP3server = "pop.elkjop.int";
            epostSMTPport = 995;
            epostSMTPssl = false;

            serviceAutoImportMinutter = 60;
            serviceAutoImportFraIndex = 8;
            serviceAutoImportTilIndex = 20;
            serviceShowTrend = true;
            serviceEgenServiceFilter = "elkjøp";

            sortIndex = 1;
            macroElguide = @"C:\Program Files\Eric's TelNet98\no_trondheim.ETN";
            kravHitrateData = 0.3M;
            kravHitrateAudioVideo = 0.1M;
            kravHitrateTele = 0.1M;
            kravHitrateNettbrett = 0.1M;
            epostEmne = "Daglig tjeneste ranking (ukedag) (dag) (måned) (autogenerert)";
            epostEmneQuick = "Kveldstall (ukedag) (dag) (måned) (autogenerert)";
            epostHour = 8;
            epostHourQuick = 20;
            epostHourQuickSaturday = 18;
            epostHourQuickSunday = 18;
            ignoreSunday = true;
            visningNull = "&nbsp;";
            graphVis = true;
            graphLimit = 1;
            kolInntjen = true;
            kolSalgspris = false;
            kolForkort = true;
            kolForkortLengde = 9;
            graphButikk = false;
            graphData = true;
            graphAudioVideo = true;
            graphTele = true;
            graphOversikt = false;
            graphDager = 30;
            graphResX = 2000;
            graphResY = 600;
            graphWidth = 940;
            favAvdeling = "";
            savedPage = "Data";
            pdfLandscape = false;
            bestofCompareIndex = 0;
            pdfZoom = 1;
            graphScreenDPI = 0.75f;
            chainElkjop = true;
            serviceHistoryDays = 60;
        }

        public int[] GetMainproductGroups(int sector)
        {
            switch (sector)
            {
                case 0:
                    List<int> list = new List<int>();
                    list.AddRange(mainproductListMda);
                    list.AddRange(mainproductListAv);
                    list.AddRange(mainproductListSda);
                    list.AddRange(mainproductListTele);
                    list.AddRange(mainproductListData);
                    return list.ToArray();
                case 1:
                    return mainproductListMda;
                case 2:
                    return mainproductListAv;
                case 3:
                    return mainproductListSda;
                case 4:
                    return mainproductListTele;
                case 5:
                    return mainproductListData;
                case 6:
                    return new int[1] { 0 };
                case 7:
                    List<int> list7 = new List<int>();
                    list7.AddRange(mainproductListMda);
                    list7.AddRange(mainproductListAv);
                    list7.AddRange(mainproductListSda);
                    list7.AddRange(mainproductListTele);
                    list7.AddRange(mainproductListData);
                    return list7.ToArray();
                case 8:
                    List<int> list8 = new List<int>();
                    list8.AddRange(mainproductListMda);
                    list8.AddRange(mainproductListAv);
                    list8.AddRange(mainproductListSda);
                    list8.AddRange(mainproductListTele);
                    list8.AddRange(mainproductListData);
                    return list8.ToArray();
                case 9:
                    List<int> list9 = new List<int>();
                    list9.AddRange(mainproductListMda);
                    list9.AddRange(mainproductListAv);
                    list9.AddRange(mainproductListSda);
                    list9.AddRange(mainproductListTele);
                    list9.AddRange(mainproductListData);
                    return list9.ToArray();
                default:
                    return new int[1] { 0 };
            }
        }

        public int[] GetAccessorieGroups(int sector)
        {
            switch (sector)
            {
                case 0:
                    List<int> list = new List<int>();
                    list.AddRange(accessoriesListMda);
                    list.AddRange(accessoriesListAv);
                    list.AddRange(accessoriesListSda);
                    list.AddRange(accessoriesListTele);
                    list.AddRange(accessoriesListData);
                    return list.ToArray();
                case 1:
                    return accessoriesListMda;
                case 2:
                    return accessoriesListAv;
                case 3:
                    return accessoriesListSda;
                case 4:
                    return accessoriesListTele;
                case 5:
                    return accessoriesListData;
                case 6:
                    return new int[1] { 0 };
                case 7:
                    List<int> list7 = new List<int>();
                    list7.AddRange(accessoriesListMda);
                    list7.AddRange(accessoriesListAv);
                    list7.AddRange(accessoriesListSda);
                    list7.AddRange(accessoriesListTele);
                    list7.AddRange(accessoriesListData);
                    return list7.ToArray();
                case 8:
                    List<int> list8 = new List<int>();
                    list8.AddRange(accessoriesListMda);
                    list8.AddRange(accessoriesListAv);
                    list8.AddRange(accessoriesListSda);
                    list8.AddRange(accessoriesListTele);
                    list8.AddRange(accessoriesListData);
                    return list8.ToArray();
                case 9:
                    List<int> list9 = new List<int>();
                    list9.AddRange(accessoriesListMda);
                    list9.AddRange(accessoriesListAv);
                    list9.AddRange(accessoriesListSda);
                    list9.AddRange(accessoriesListTele);
                    list9.AddRange(accessoriesListData);
                    return list9.ToArray();
                default:
                    return new int[1] { 0 };
            }
        }

        public void ResetVarekoder(List<VarekodeList> vareList)
        {
            var item = new VarekodeList();
            item.Insert("KG00", 0, 100, 695, "Data", true, "RTG00", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTG00", 0, 100, 699, "Data", true, "RTG00", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAB06", 0, 0, 995, "Data", false, "SAB06", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SA06", 0, 0, 1295, "Data", false, "SA06", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAB12", 37.5M, 200, 1395, "Data", true, "SAB12", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SA12", 75, 200, 1699, "Data", true, "SA12", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAB24", 50, 300, 2195, "Data", true, "SAB24", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SA24", 100, 300, 2499, "Data", true, "SA24", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAB36", 75, 400, 2995, "Data", true, "SAB36", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SA36", 200, 400, 3299, "Data", true, "SA36", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGIPAD", 0, 0, 499, "Nettbrett", true, "RTGIPAD", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGTABLET", 0, 0, 499, "Nettbrett", true, "RTGTABLET", true);
            vareList.Add(item);

            item = new VarekodeList();
            item.Insert("SATABLET06", 0, 0, 699, "Nettbrett", true, "SATABLET", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SATABLET12", 0, 0, 995, "Nettbrett", true, "SATABLET", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SATABLET24", 0, 0, 1495, "Nettbrett", true, "SATABLET", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SATABLET36", 0, 0, 1995, "Nettbrett", true, "SATABLET", true);
            vareList.Add(item);

            item = new VarekodeList();
            item.Insert("RTGMOBILE", 0, 0, 299, "Tele", true, "RTGMOBILE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAMOBILE06", 0, 0, 499, "Tele", true, "SAMOBILE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAMOBILE12", 0, 0, 699, "Tele", true, "SAMOBILE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAMOBILE24", 0, 0, 999, "Tele", true, "SAMOBILE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAMOBILE36", 0, 0, 1295, "Tele", true, "SAMOBILE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGMOBILESAFE", 0, 0, 499, "Tele", true, "RTGMOBILESAFE", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("MOBSECURITY", 0, 0, 249, "Tele", true, "MOBSECURITY", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGGPS", 0, 0, 295, "Tele", false, "RTGGPS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGTV", 0, 0, 995, "AudioVideo", true, "RTGTV", true);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RTGREMOTE", 0, 0, 295, "AudioVideo", true, "RTGREMOTE", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("TIMEPRIS15MIN", 0, 0, 200, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("TIMEPRIS30MIN", 0, 0, 395, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("TIMEPRIS60MIN", 0, 0, 595, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);

            item = new VarekodeList();
            item.Insert("MANHOUR15MIN", 0, 0, 199, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("MANHOUR30MIN", 0, 0, 395, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("MANHOUR60MIN", 0, 0, 595, "Data", false, "TIMEPRIS", false);
            vareList.Add(item);

            item = new VarekodeList();
            item.Insert("HEALTHCHECK", 0, 0, 0, "Data", false, "HEALTHCHECK", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SERVICEBACKUP", 0, 0, 0, "Data", false, "SERVICEBACKUP", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("BACKUP", 0, 0, 1244, "Data", false, "BACKUP", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("RECHDD", 0, 0, 0, "Nettbrett", false, "RECHDD", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("CLOUDSMALL", 0, 0, 195, "Data", false, "CLOUDSMALL", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("CLOUDUNLIM", 0, 0, 285, "Data", false, "CLOUDUNLIM", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("OSUNLIM", 0, 0, 285, "Data", false, "OSUNLIM", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("INGENFEIL", 0, 0, 0, "Data", false, "INGENFEIL", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAGPS36", 0, 0, 599, "Tele", true, "SAGPS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAGPS24", 0, 0, 499, "Tele", true, "SAGPS", false);
            vareList.Add(item);
            item = new VarekodeList();
            item.Insert("SAGPS12", 0, 0, 399, "Tele", true, "SAGPS", false);
            vareList.Add(item);
        }
    }

    public class XmlColor
    {
        private Color color_ = Color.Black;

        public XmlColor() { }
        public XmlColor(Color c) { color_ = c; }


        public Color ToColor()
        {
            return color_;
        }

        public void FromColor(Color c)
        {
            color_ = c;
        }

        public static implicit operator Color(XmlColor x)
        {
            return x.ToColor();
        }

        public static implicit operator XmlColor(Color c)
        {
            return new XmlColor(c);
        }

        [XmlAttribute]
        public string Web
        {
            get { return ColorTranslator.ToHtml(color_); }
            set
            {
                try
                {
                    if (Alpha == 0xFF) // preserve named color value if possible
                        color_ = ColorTranslator.FromHtml(value);
                    else
                        color_ = Color.FromArgb(Alpha, ColorTranslator.FromHtml(value));
                }
                catch (Exception)
                {
                    color_ = Color.Black;
                }
            }
        }

        [XmlAttribute]
        public byte Alpha
        {
            get { return color_.A; }
            set
            {
                if (value != color_.A) // avoid hammering named color if no alpha change
                    color_ = Color.FromArgb(value, color_);
            }
        }

        public bool ShouldSerializeAlpha() { return Alpha < 0xFF; }
    }
}
