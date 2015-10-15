using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using KGSA.Properties;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;

namespace KGSA
{
    public partial class FormSettings : Form
    {
        private BackgroundWorker bwUnlockTest = new BackgroundWorker();
        private readonly Timer timerMsgClear = new Timer();
        FormMain main;
        public bool forceUpdate = false;
        public FormSettings(FormMain form)
        {
            this.main = form;
            InitializeComponent();
            timerMsgClear.Tick += timer;
            Logg.Debug("Innstillinger åpnet.");
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            ImportAppSettings();
            DBsize();
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

        static String BytesToString(long byteCount)
        {
            string[] suf = { " B", " KB", " MB", " GB", " TB", " PB", " EB" }; //Longs run out around EB
            if (byteCount == 0)
                return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString() + suf[place];
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

        private void ImportAppSettings()
        {
            // Fane: GENERELT
            // GENERELT - Elguide
            if (main.appConfig.Avdeling >= 1000 && main.appConfig.Avdeling <= 1700)
            {
                textBoxAvd.BackColor = ColorTranslator.FromHtml("#b9dc9c");
                textBoxAvd.Text = main.appConfig.Avdeling.ToString();
                button1.Enabled = true;
            }
            else
            {
                textBoxAvd.BackColor = ColorTranslator.FromHtml("#ecc3ad");
                textBoxAvd.Text = "";
                button1.Enabled = false;
            }
            textBoxCSVexportFolder.Text = main.appConfig.csvElguideExportFolder;

            // GENERELT - Import
            if (main.appConfig.importSetting == "Normal")
                comboBoxImport.SelectedIndex = 0;
            if (main.appConfig.importSetting == "Full")
                comboBoxImport.SelectedIndex = 1;
            if (main.appConfig.importSetting == "FullFavoritt")
                comboBoxImport.SelectedIndex = 2;

            // GENERELT - Program
            checkMinimizeToTray.Checked = main.appConfig.WindowMinimizeToTray;
            checkExitToTray.Checked = main.appConfig.WindowExitToTray;
            checkBoxActivateSqlCache.Checked = main.appConfig.useSqlCache;
            checkBoxAdvanced.Checked = main.appConfig.showAdvancedFunctions;
            if (main.appConfig.showAdvancedFunctions)
                AddAdvancedTabs();
            else
                RemoveAdvancedTabs();
            checkBoxActivateExperimental.Checked = main.appConfig.experimental;
            if (main.appConfig.experimental)
                AddExperimentalTabs();
            else
                RemoveExperimentalTabs();

            // Fanse: RANKING
            // RANKING - Utseende
            checkBoxRankUtseendeInntjen.Checked = main.appConfig.kolInntjen;
            checkVisProv.Checked = main.appConfig.kolProv;
            kolOmsetningCheckBox.Checked = main.appConfig.kolSalgspris;
            kolRabattCheckBox.Checked = main.appConfig.kolRabatt;
            checkBoxKombkol.Checked = !main.appConfig.kolVarekoder;
            if (main.appConfig.sortIndex > -1)
                comboBoxRankingSort.SelectedIndex = main.appConfig.sortIndex;
            checkRankAvdShowAll.Checked = main.appConfig.rankingAvdelingShowAll;

            // RANKING - Vinnprodukter
            checkBoxVinn.Checked = main.appConfig.vinnDatoFraTil;
            if (main.appConfig.vinnFrom > FormMain.rangeMin)
                dateTimeVinnFrom.Value = main.appConfig.vinnFrom;
            if (main.appConfig.vinnTo > FormMain.rangeMin)
            dateTimeVinnTo.Value = main.appConfig.vinnTo;
            checkBoxVinnSisteDagVare.Checked = main.appConfig.vinnSisteDagVare;
            checkBoxVinnVisVarekoder.Checked = main.appConfig.vinnVisVarekoder;
            checkBoxVinnVisVarekoderEkstra.Checked = main.appConfig.vinnVisVarekoderExtra;
            checkBoxVinnSimple.Checked = main.appConfig.vinnEnkelModus;
            checkBoxVinnVisAftersales.Checked = main.appConfig.vinnVisAftersales;

            // RANKING - Hitrate krav
            checkKravNettbrett.Checked = main.appConfig.kolKravNettbrett;
            checkKravData.Checked = main.appConfig.kolKravData;
            checkKravAudioVideo.Checked = main.appConfig.kolKravAudioVideo;
            checkKravTele.Checked = main.appConfig.kolKravTele;
            SetNumericValue(numericKravHitrateData, main.appConfig.kravHitrateData);
            SetNumericValue(numericKravHitrateAudioVideo, main.appConfig.kravHitrateAudioVideo);
            SetNumericValue(numericKravHitrateTele, main.appConfig.kravHitrateTele);
            SetNumericValue(numericKravHitrateNettbrett, main.appConfig.kravHitrateNettbrett);

            // RANKING - Valg
            checkBoxIgnoreSunday.Checked = main.appConfig.ignoreSunday;
            checkBoxRankingShowLastWeek.Checked = main.appConfig.rankingShowLastWeek;
            comboBoxRankingFjoraaret.SelectedIndex = main.appConfig.rankingCompareLastyear;
            comboBoxRankingForrigeMaaned.SelectedIndex = main.appConfig.rankingCompareLastmonth;

            // Fane: PDF ----------------------------------------------------------------------------
            // PDF - PDF Export
            
            checkBoxPDFbutikk.Checked = main.appConfig.pdfVisButikk;
            checkBoxPDFknowhow.Checked = main.appConfig.pdfVisKnowHow;
            checkBoxPDFdata.Checked = main.appConfig.pdfVisData;
            checkBoxPDFaudiovideo.Checked = main.appConfig.pdfVisAudioVideo;
            checkBoxPDFtele.Checked = main.appConfig.pdfVisTele;
            checkBoxPDFlister.Checked = main.appConfig.pdfVisLister;
            checkBoxPDFoversikt.Checked = main.appConfig.pdfVisOversikt;
            checkBoxPDFtoppselgere.Checked = main.appConfig.pdfVisToppselgere;
            checkBoxPDFvinnprodukter.Checked = main.appConfig.pdfVisVinnprodukter;

            checkBoxPDFtjenester.Checked = main.appConfig.pdfVisTjenester;
            checkBoxPDFsnittpriser.Checked = main.appConfig.pdfVisSnittpriser;

            checkBoxPdfVisService.Checked = main.appConfig.pdfVisService;

            checkBoxPDFobsolete.Checked = main.appConfig.pdfVisObsolete;
            checkBoxPDFobsoleteList.Checked = main.appConfig.pdfVisObsoleteList;
            checkBoxPDFweekly.Checked = main.appConfig.pdfVisWeekly;
            checkBoxPDFprisguide.Checked = main.appConfig.pdfVisPrisguide;

            checkBoxPdfVisBudget.Checked = main.appConfig.pdfVisBudsjett;

            // PDF - Format
            if (main.appConfig.pdfLandscape)
                radioButtonPDForiantationSide.Checked = true;
            else
                radioButtonPDForiantationUp.Checked = true;
            SetNumericValue(numericPDFzoon, main.appConfig.pdfZoom);
            checkBoxPdfUtvidetGrafer.Checked = main.appConfig.pdfExpandedGraphs;
            checkBoxPdfToc.Checked = main.appConfig.pdfTableOfContent;

            // PDF - PDF Export
            textBoxExportPDF.Text = main.appConfig.pdfExportFolder;
            checkPDFExport.Checked = main.appConfig.pdfExport;

            // Fane: TOPPSELGERE --------------------------------------------------------------------------
            // TOPPSELGERE - Sortering
            checkBoxToppselgereSortTjenesterInntjen.Checked = main.appConfig.bestofSortTjenesterSecondary;
            checkBoxToppselgereSortInntjening.Checked = main.appConfig.bestofSortInntjenSecondary;

            // TOPPSELGERE - Antall
            comboBoxAntallInntjen.SelectedIndex = main.appConfig.bestofTallInntjen;
            comboBoxAntallTA.SelectedIndex = main.appConfig.bestofTallTA;
            comboBoxAntallStrom.SelectedIndex = main.appConfig.bestofTallStrom;
            comboBoxAntallRTG.SelectedIndex = main.appConfig.bestofTallTjen;
            comboBoxAntallFinans.SelectedIndex = main.appConfig.bestofTallFinans;

            // TOPPSELGERE - Annet
            checkBoxBestofCompareChange.Checked = main.appConfig.bestofCompareChange;
            checkBoxToppselgereVisBesteFjoråret.Checked = main.appConfig.bestofVisBesteLastYear;
            if (main.appConfig.bestofVisBesteLastYearTotal)
                radioButtonToppselgereVisBesteFjoråretTotalt.Checked = true;
            else
                radioButtonToppselgereVisBesteFjoråretMtd.Checked = true;
            checkBoxTopplisteVisBesteSisteåpningsdag.Checked = main.appConfig.bestofVisBesteLastOpenday;
            checkBoxBestofHoppoverKasse.Checked = main.appConfig.bestofHoppoverKasse;

            // TOPPSELGERE - Lister
            checkBoxListerInntjening.Checked = main.appConfig.listerVisInntjen;
            checkBoxListerAccessories.Checked = main.appConfig.listerVisAccessories;
            SetNumericValue(numericListerMax, main.appConfig.listerMaxLinjer);

            // Fane: OVERSIKT ---------------------------------------------------------------------------
            // OVERSIKT - Generelt
            checkBoxOversiktBegrens.Checked = main.appConfig.oversiktFilterToDepartments;
            checkBoxOversiktHideAftersales.Checked = main.appConfig.oversiktHideAftersales;
            checkBoxOversiktHideKitchen.Checked = main.appConfig.oversiktHideKitchen;

            // OVERSIKT - Krav
            checkOversiktVisKrav.Checked = main.appConfig.oversiktKravVis;
            checkKravFinans.Checked = main.appConfig.oversiktKravFinans;
            checkKravMod.Checked = main.appConfig.oversiktKravMod;
            checkKravStrom.Checked = main.appConfig.oversiktKravStrom;
            checkKravRtgsa.Checked = main.appConfig.oversiktKravRtgsa;
            checkKravAntallFinans.Checked = main.appConfig.oversiktKravFinansAntall;
            checkKravAntallMod.Checked = main.appConfig.oversiktKravModAntall;
            checkKravAntallStrom.Checked = main.appConfig.oversiktKravStromAntall;
            checkKravAntallRtgsa.Checked = main.appConfig.oversiktKravRtgsaAntall;
            checkKravMTD.Checked = main.appConfig.oversiktKravMtd;
            checkKravMtdShowTarget.Checked = main.appConfig.oversiktKravMtdShowTarget;

            // Fane: UTSEENDE -----------------------------------------------------------------
            // UTSEENDE - Data hitrate farger
            color1.BackColor = main.appConfig.color1;
            if (main.appConfig.color1inv) { color1.ForeColor = Color.White; }
            numericUpDown01.Value = main.appConfig.color1max;
            color2.BackColor = main.appConfig.color2;
            if (main.appConfig.color2inv) { color2.ForeColor = Color.White; }
            numericUpDown02.Value = main.appConfig.color2min;
            numericUpDown03.Value = main.appConfig.color2max;
            color3.BackColor = main.appConfig.color3;
            if (main.appConfig.color3inv) { color3.ForeColor = Color.White; }
            numericUpDown04.Value = main.appConfig.color3min;
            numericUpDown05.Value = main.appConfig.color3max;
            color4.BackColor = main.appConfig.color4;
            if (main.appConfig.color4inv) { color4.ForeColor = Color.White; }
            numericUpDown06.Value = main.appConfig.color4min;
            numericUpDown07.Value = main.appConfig.color4max;
            color5.BackColor = main.appConfig.color5;
            if (main.appConfig.color5inv) { color5.ForeColor = Color.White; }
            numericUpDown08.Value = main.appConfig.color5min;
            numericUpDown09.Value = main.appConfig.color5max;
            color6.BackColor = main.appConfig.color6;
            if (main.appConfig.color6inv) { color6.ForeColor = Color.White; }
            numericUpDown10.Value = main.appConfig.color6min;

            // UTSEENDE - Nettbrett farger
            ncolor1.BackColor = main.appConfig.ncolor1;
            if (main.appConfig.ncolor1inv) { ncolor1.ForeColor = Color.White; }
            numericUpDown1.Value = main.appConfig.ncolor1max;
            ncolor2.BackColor = main.appConfig.ncolor2;
            if (main.appConfig.ncolor2inv) { ncolor2.ForeColor = Color.White; }
            numericUpDown9.Value = main.appConfig.ncolor2min;
            numericUpDown11.Value = main.appConfig.ncolor2max;
            ncolor3.BackColor = main.appConfig.ncolor3;
            if (main.appConfig.ncolor3inv) { ncolor3.ForeColor = Color.White; }
            numericUpDown8.Value = main.appConfig.ncolor3min;
            numericUpDown7.Value = main.appConfig.ncolor3max;
            ncolor4.BackColor = main.appConfig.ncolor4;
            if (main.appConfig.ncolor4inv) { ncolor4.ForeColor = Color.White; }
            numericUpDown6.Value = main.appConfig.ncolor4min;
            numericUpDown5.Value = main.appConfig.ncolor4max;
            ncolor5.BackColor = main.appConfig.ncolor5;
            if (main.appConfig.ncolor5inv) { ncolor5.ForeColor = Color.White; }
            numericUpDown4.Value = main.appConfig.ncolor5min;
            numericUpDown3.Value = main.appConfig.ncolor5max;
            ncolor6.BackColor = main.appConfig.ncolor6;
            if (main.appConfig.ncolor6inv) { ncolor6.ForeColor = Color.White; }
            numericUpDown2.Value = main.appConfig.ncolor6min;

            // UTSEENDE - Format og logo
            if (main.appConfig.visningNull == "&nbsp;")
                visningNullTextBox.Text = "";
            else
                visningNullTextBox.Text = main.appConfig.visningNull;
            checkBoxJevnfarge.Checked = main.appConfig.visningJevnfarge;
            checkBoxTruncate.Checked = main.appConfig.kolForkort;
            if (main.appConfig.chainElkjop)
            {
                panelPictureElkjop.BackColor = Color.Black;
                panelPictureLefdal.BackColor = Color.White;
            }
            else
            {
                panelPictureElkjop.BackColor = Color.White;
                panelPictureLefdal.BackColor = Color.Black;
            }

            // Fane: FAVORITTER
            // FAVORITTER - Favoritt avdelinger
            if (main.appConfig.favAvdeling.Length > 3)
            {
                string[] avdStr = main.appConfig.favAvdeling.Split(',');
                Avdeling avdObject = new Avdeling();
                foreach (string avd in avdStr)
                {
                    string str = avdObject.Get(avd);
                    if (str.Length == 4)
                        listBoxFavAvd.Items.Add(avd);
                    else
                        listBoxFavAvd.Items.Add(avd + ": " + str);
                }
            }
            checkBoxFav.Checked = main.appConfig.favVis;

            // Fane: AVANSERT
            // AVANSERT - hovedprodukt varegrupper
            if (VerifyIntArrayToString(main.appConfig.mainproductListMda))
                textBoxHovedMda.Text = IntArrayToString(main.appConfig.mainproductListMda);
            else
                textBoxHovedMda.Text = "131,132,134,135,136,143,144,145";

            if (VerifyIntArrayToString(main.appConfig.mainproductListAv))
                textBoxHovedAv.Text = IntArrayToString(main.appConfig.mainproductListAv);
            else
                textBoxHovedAv.Text = "224,273";

            if (VerifyIntArrayToString(main.appConfig.mainproductListSda))
                textBoxHovedSda.Text = IntArrayToString(main.appConfig.mainproductListSda);
            else
                textBoxHovedSda.Text = "301,346";

            if (VerifyIntArrayToString(main.appConfig.mainproductListTele))
                textBoxHovedTele.Text = IntArrayToString(main.appConfig.mainproductListTele);
            else
                textBoxHovedTele.Text = "431,447";

            if (VerifyIntArrayToString(main.appConfig.mainproductListData))
                textBoxHovedData.Text = IntArrayToString(main.appConfig.mainproductListData);
            else
                textBoxHovedData.Text = "531,533,534";

            // AVANSERT - tilbehør varegrupper
            if (VerifyIntArrayToString(main.appConfig.accessoriesListMda))
                textBoxAccMda.Text = IntArrayToString(main.appConfig.accessoriesListMda);
            else
                textBoxAccMda.Text = "195";

            if (VerifyIntArrayToString(main.appConfig.accessoriesListAv))
                textBoxAccAv.Text = IntArrayToString(main.appConfig.accessoriesListAv);
            else
                textBoxAccAv.Text = "214";

            if (VerifyIntArrayToString(main.appConfig.accessoriesListSda))
                textBoxAccSda.Text = IntArrayToString(main.appConfig.accessoriesListSda);
            else
                textBoxAccSda.Text = "395";

            if (VerifyIntArrayToString(main.appConfig.accessoriesListTele))
                textBoxAccTele.Text = IntArrayToString(main.appConfig.accessoriesListTele);
            else
                textBoxAccTele.Text = "487";

            if (VerifyIntArrayToString(main.appConfig.accessoriesListData))
                textBoxAccData.Text = IntArrayToString(main.appConfig.accessoriesListData);
            else
                textBoxAccData.Text = "552,569,589";

            // Fane: E-POST
            // E-POST - Server
            textBoxEpostSMTPserver.Text = main.appConfig.epostSMTPserver;
            textBoxEpostSMTPport.Text = main.appConfig.epostSMTPport.ToString();
            checkBoxEmailUseBcc.Checked = main.appConfig.epostBrukBcc;
            checkBoxEmailUseSsl.Checked = main.appConfig.epostSMTPssl;
            textBoxEpostAvsender.Text = main.appConfig.epostAvsender;
            textBoxEpostAvsenderNavn.Text = main.appConfig.epostAvsenderNavn;
            textBoxEpostPOP3server.Text = main.appConfig.epostPOP3server;
            textBoxEpostPOP3port.Text = main.appConfig.epostPOP3port.ToString();
            checkBoxPOP3UseSsl.Checked = main.appConfig.epostPOP3ssl;
            textBoxEpostPOP3bruker.Text = main.appConfig.epostPOP3username;
            var aes = new SimpleAES();
            if (main.appConfig.epostPOP3password != "")
                textBoxEpostPOP3pass.Text = aes.DecryptString(main.appConfig.epostPOP3password);
            else
                textBoxEpostPOP3pass.Text = "";
            SetNumericValue(numericEpostPOP3searchLimit, main.appConfig.epostPOP3searchLimit);

            // E-POST - Annet
            textBoxEpostNesteMelding.Text = main.appConfig.epostNesteMelding.Replace("\n", Environment.NewLine);


            // Fane: SERVICE
            // SERVICE - Service generelt
            textServiceEgenServiceFilter.Text = main.appConfig.serviceEgenServiceFilter;
            checkServiceShowHistory.Checked = main.appConfig.serviceShowHistory;
            checkServiceShowHistoryGraph.Checked = main.appConfig.serviceShowHistoryGraph;
            checkServiceShowTrend.Checked = main.appConfig.serviceShowTrend;
            SetNumericValue(numericServiceHistoryDays, main.appConfig.serviceHistoryDays);
            checkBoxServiceFerdigService.Checked = main.appConfig.serviceFerdigServiceStats;
            SetNumericValue(numericServiceFerdigStatsAntall, main.appConfig.serviceFerdigServiceStatsAntall);

            // Fane: GRAF ----------------------------------------------------------------------------
            // GRAF - Grafikk
            checkBoxVisGrafikk.Checked = main.appConfig.graphVis;
            SetNumericValue(numericGraphDPI, main.appConfig.graphScreenDPI);
            textBoxGraphWidth.Text = main.appConfig.graphWidth.ToString();
            checkBoxGrafikkExtra.Checked = main.appConfig.graphExtra;
            checkBoxGrafikkAdvanced.Checked = main.appConfig.graphAdvanced;
            checkBoxGrafikkButikk.Checked = main.appConfig.graphButikk;
            checkBoxGrafikkKnowHow.Checked = main.appConfig.graphKnowHow;
            checkBoxGrafikkData.Checked = main.appConfig.graphData;
            checkBoxGrafikkAV.Checked = main.appConfig.graphAudioVideo;
            checkBoxGrafikkTele.Checked = main.appConfig.graphTele;
            checkBoxGrafikkOversikt.Checked = main.appConfig.graphOversikt;
            textBoxGraphAntall.Text = main.appConfig.graphDager.ToString();
            if (main.appConfig.graphResX > 900 && main.appConfig.graphResY > 300)
            {
                textBoxGraphX.Text = main.appConfig.graphResX.ToString();
                textBoxGraphY.Text = main.appConfig.graphResY.ToString();
            }
            else
            {
                textBoxGraphX.Text = "900";
                textBoxGraphY.Text = "300";
            }

            // Fane: LAGER ------------------------------------------------------------------------------------
            // LAGER - Generelt
            if (main.appConfig.storeCompareMtd)
                radioLagerUtviklingMTD.Checked = true;
            else
                radioLagerUtviklingDato.Checked = true;
            if (main.appConfig.dbStoreViewpoint.Date < main.appConfig.dbStoreTo && main.appConfig.dbStoreViewpoint > main.appConfig.dbStoreFrom)
                pickerLagerViewpoint.Value = main.appConfig.dbStoreViewpoint;
            else if (main.appConfig.dbStoreFrom > FormMain.rangeMin)
                pickerLagerViewpoint.Value = main.appConfig.dbStoreFrom;
            else
                pickerLagerViewpoint.Value = DateTime.Now;
            checkBoxStoreShowStoreTwo.Checked = main.appConfig.storeShowStoreTwo;
            SetNumericValue(numericUpDownLagerPriserMaxDager, main.appConfig.storeMaxAgePrizes);

            // LAGER - Ukurante varer
            SetNumericValue(numericStoreObsoleteFilter, main.appConfig.storeObsoleteFilterPercent);
            SetNumericValue(numericStoreObsoleteFilterMax, main.appConfig.storeObsoleteFilterMax);

            checkLagerHovedprodukter.Checked = main.appConfig.storeObsoleteListMainProductsOnly;
            checkBoxObsoleteFilterMainStoreOnly.Checked = main.appConfig.storeObsoleteFilterMainStoreOnly;

            comboBoxStoreObsoleteSortBy.Text = main.appConfig.storeObsoleteSortBy;
            checkBoxStoreObsoleteSortAsc.Checked = main.appConfig.storeObsoleteSortAsc;

            // LAGER - Prisguide.no og Ukenytt
            SetNumericValue(numericUpDownLagerPrisguideImportPages, main.appConfig.onlinePrisguidePagesToImport);

            // Fane: WEBSERVER --------------------------------------------------------------------------------
            // WEBSERVER - Generelt
            checkBoxWebserverEnabled.Checked = main.appConfig.webserverEnabled;
            SetNumericValue(numericWebserverPort, main.appConfig.webserverPort);
            comboBoxWebserverBindings.Items.Clear();
            List<string> bindings = GetAllBindings();
            if (bindings != null)
            {
                foreach (string address in bindings)
                    comboBoxWebserverBindings.Items.Add(address);
            }
            comboBoxWebserverBindings.Text = main.appConfig.webserverHost;

            // WEBSERVER - Sikkerhet
            checkBoxWebserverSimpleAut.Checked = main.appConfig.webserverRequireSimpleAuthentication;
            textBoxWebserverUser.Text = main.appConfig.webserverUser;
            if (main.appConfig.webserverPassword != "")
                textBoxWebserverPassword.Text = aes.DecryptString(main.appConfig.webserverPassword);
            else
                textBoxWebserverPassword.Text = "";

            // Fane: VEDLIKEHOLD ------------------------------------------------------------------------------
            // VEDLIKEHOLD - Annet
            debugCheckBox.Checked = main.appConfig.debug;
        }

        bool IsValidEmail(string email)
        {
            try
            {
                if (email.Length < 3)
                    return false;
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private bool ExportSettings()
        {
            try
            {
                // Fane: GENERELT -------------------------------------------------------------------------
                // GENERELT - Elguide
                if (textBoxAvd.Text.Length == 4)
                    main.appConfig.Avdeling = Convert.ToInt32(textBoxAvd.Text);
                main.appConfig.csvElguideExportFolder = textBoxCSVexportFolder.Text;
                
                // GENERELT - Import
                if (comboBoxImport.SelectedIndex == 0)
                    main.appConfig.importSetting = "Normal";
                if (comboBoxImport.SelectedIndex == 1)
                    main.appConfig.importSetting = "Full";
                if (comboBoxImport.SelectedIndex == 2)
                    main.appConfig.importSetting = "FullFavoritt";

                // GENERELT - Program
                main.appConfig.WindowMinimizeToTray = checkMinimizeToTray.Checked;
                main.appConfig.WindowExitToTray = checkExitToTray.Checked;
                main.appConfig.useSqlCache = checkBoxActivateSqlCache.Checked;
                main.appConfig.showAdvancedFunctions = checkBoxAdvanced.Checked;
                main.appConfig.experimental = checkBoxActivateExperimental.Checked;

                // Fane: RANKING --------------------------------------------------------------------------
                // RANKING - Utseende
                main.appConfig.kolInntjen = checkBoxRankUtseendeInntjen.Checked;
                main.appConfig.kolProv = checkVisProv.Checked;
                main.appConfig.kolRabatt = kolRabattCheckBox.Checked;
                main.appConfig.kolSalgspris = kolOmsetningCheckBox.Checked;
                main.appConfig.kolVarekoder = !checkBoxKombkol.Checked;
                main.appConfig.sortIndex = Convert.ToInt32(comboBoxRankingSort.SelectedIndex.ToString());
                main.appConfig.rankingAvdelingShowAll = checkRankAvdShowAll.Checked;

                // RANKING - Hitrate krav
                main.appConfig.kravHitrateNettbrett = numericKravHitrateNettbrett.Value;
                main.appConfig.kravHitrateData = numericKravHitrateData.Value;
                main.appConfig.kravHitrateAudioVideo = numericKravHitrateAudioVideo.Value;
                main.appConfig.kravHitrateTele = numericKravHitrateTele.Value;
                main.appConfig.kolKravNettbrett = checkKravNettbrett.Checked;
                main.appConfig.kolKravData = checkKravData.Checked;
                main.appConfig.kolKravAudioVideo = checkKravAudioVideo.Checked;
                main.appConfig.kolKravTele = checkKravTele.Checked;

                // RANKING - Valg
                main.appConfig.ignoreSunday = checkBoxIgnoreSunday.Checked;
                main.appConfig.rankingShowLastWeek = checkBoxRankingShowLastWeek.Checked;
                main.appConfig.rankingCompareLastyear = comboBoxRankingFjoraaret.SelectedIndex;
                main.appConfig.rankingCompareLastmonth = comboBoxRankingForrigeMaaned.SelectedIndex;

                // RANKING - Vinnprodukter
                main.appConfig.vinnDatoFraTil = checkBoxVinn.Checked;
                main.appConfig.vinnFrom = dateTimeVinnFrom.Value;
                main.appConfig.vinnTo = dateTimeVinnTo.Value;
                main.appConfig.vinnSisteDagVare = checkBoxVinnSisteDagVare.Checked;
                main.appConfig.vinnVisVarekoder = checkBoxVinnVisVarekoder.Checked;
                main.appConfig.vinnVisVarekoderExtra = checkBoxVinnVisVarekoderEkstra.Checked;
                main.appConfig.vinnEnkelModus = checkBoxVinnSimple.Checked;
                main.appConfig.vinnVisAftersales = checkBoxVinnVisAftersales.Checked;

                // Fane: PDF ----------------------------------------------------------------------------
                // PDF - PDF Export
                main.appConfig.pdfVisOversikt = checkBoxPDFoversikt.Checked;
                main.appConfig.pdfVisToppselgere = checkBoxPDFtoppselgere.Checked;
                main.appConfig.pdfVisButikk = checkBoxPDFbutikk.Checked;
                main.appConfig.pdfVisKnowHow = checkBoxPDFknowhow.Checked;
                main.appConfig.pdfVisData = checkBoxPDFdata.Checked;
                main.appConfig.pdfVisAudioVideo = checkBoxPDFaudiovideo.Checked;
                main.appConfig.pdfVisTele = checkBoxPDFtele.Checked;
                main.appConfig.pdfVisLister = checkBoxPDFlister.Checked;
                main.appConfig.pdfVisVinnprodukter = checkBoxPDFvinnprodukter.Checked;

                main.appConfig.pdfVisTjenester = checkBoxPDFtjenester.Checked;
                main.appConfig.pdfVisSnittpriser = checkBoxPDFsnittpriser.Checked;

                main.appConfig.pdfVisObsolete = checkBoxPDFobsolete.Checked;
                main.appConfig.pdfVisObsoleteList = checkBoxPDFobsoleteList.Checked;
                main.appConfig.pdfVisWeekly = checkBoxPDFweekly.Checked;
                main.appConfig.pdfVisPrisguide = checkBoxPDFprisguide.Checked;

                main.appConfig.pdfVisService = checkBoxPdfVisService.Checked;

                main.appConfig.pdfVisBudsjett = checkBoxPdfVisBudget.Checked;

                // PDF - Format
                main.appConfig.pdfLandscape = radioButtonPDForiantationSide.Checked;
                main.appConfig.pdfZoom = numericPDFzoon.Value;
                main.appConfig.pdfExpandedGraphs = checkBoxPdfUtvidetGrafer.Checked;
                main.appConfig.pdfTableOfContent = checkBoxPdfToc.Checked;

                // PDF - PDF Export
                main.appConfig.pdfExportFolder = textBoxExportPDF.Text;
                main.appConfig.pdfExport = checkPDFExport.Checked;

                // Fane: TOPPSELGERE -------------------------------------------------------------------
                // TOPPSELGERE - Sortering
                main.appConfig.bestofSortTjenesterSecondary = checkBoxToppselgereSortTjenesterInntjen.Checked;
                main.appConfig.bestofSortInntjenSecondary = checkBoxToppselgereSortInntjening.Checked;

                // TOPPSELGERE - Antall
                main.appConfig.bestofTallInntjen = comboBoxAntallInntjen.SelectedIndex;
                main.appConfig.bestofTallFinans = comboBoxAntallFinans.SelectedIndex;
                main.appConfig.bestofTallStrom = comboBoxAntallStrom.SelectedIndex;
                main.appConfig.bestofTallTjen = comboBoxAntallRTG.SelectedIndex;
                main.appConfig.bestofTallTA = comboBoxAntallTA.SelectedIndex;

                // TOPPSELGERE - Annet
                main.appConfig.bestofCompareChange = checkBoxBestofCompareChange.Checked;
                main.appConfig.bestofVisBesteLastYear = checkBoxToppselgereVisBesteFjoråret.Checked;
                main.appConfig.bestofVisBesteLastYearTotal = radioButtonToppselgereVisBesteFjoråretTotalt.Checked;
                main.appConfig.bestofVisBesteLastOpenday = checkBoxTopplisteVisBesteSisteåpningsdag.Checked;
                main.appConfig.bestofHoppoverKasse = checkBoxBestofHoppoverKasse.Checked;

                // TOPPSELGERE - Lister
                main.appConfig.listerVisInntjen = checkBoxListerInntjening.Checked;
                main.appConfig.listerVisAccessories = checkBoxListerAccessories.Checked;
                main.appConfig.listerMaxLinjer = Convert.ToInt32(numericListerMax.Value);

                // Fane: OVERSIKT ---------------------------------------------------------------------
                // OVERSIKT - Generelt
                main.appConfig.oversiktFilterToDepartments = checkBoxOversiktBegrens.Checked;
                main.appConfig.oversiktHideAftersales = checkBoxOversiktHideAftersales.Checked;
                main.appConfig.oversiktHideKitchen = checkBoxOversiktHideKitchen.Checked;

                // OVERSIKT - Krav
                main.appConfig.oversiktKravVis = checkOversiktVisKrav.Checked;
                main.appConfig.oversiktKravFinansAntall = checkKravAntallFinans.Checked;
                main.appConfig.oversiktKravModAntall = checkKravAntallMod.Checked;
                main.appConfig.oversiktKravStromAntall = checkKravAntallStrom.Checked;
                main.appConfig.oversiktKravRtgsaAntall = checkKravAntallRtgsa.Checked;
                main.appConfig.oversiktKravFinans = checkKravFinans.Checked;
                main.appConfig.oversiktKravMod = checkKravMod.Checked;
                main.appConfig.oversiktKravStrom = checkKravStrom.Checked;
                main.appConfig.oversiktKravRtgsa = checkKravRtgsa.Checked;
                main.appConfig.oversiktKravMtd = checkKravMTD.Checked;
                main.appConfig.oversiktKravMtdShowTarget = checkKravMtdShowTarget.Checked;

                // OVERSIKT - Format og logo
                if (visningNullTextBox.Text == "")
                    main.appConfig.visningNull = "&nbsp;";
                else
                    main.appConfig.visningNull = visningNullTextBox.Text;
                main.appConfig.visningJevnfarge = checkBoxJevnfarge.Checked;
                main.appConfig.kolForkort = checkBoxTruncate.Checked;

                // Fane: FAVORITTER --------------------------------------------------------------------------
                // FAVORITTER - Favoritter
                main.appConfig.favAvdeling = "";
                for (int i = listBoxFavAvd.Items.Count; i-- > 0; )
                    if (listBoxFavAvd.Items[i].ToString() == "")
                        listBoxFavAvd.Items.RemoveAt(i);
                for (int i = 0; i < listBoxFavAvd.Items.Count; i++)
                {
                    if (listBoxFavAvd.Items[i].ToString().StartsWith(main.appConfig.Avdeling.ToString()))
                        continue;
                    if (listBoxFavAvd.Items[i].ToString() != "")
                        main.appConfig.favAvdeling += listBoxFavAvd.Items[i].ToString().Substring(0, 4);
                    if ((i + 1) < listBoxFavAvd.Items.Count)
                        main.appConfig.favAvdeling += ",";
                }
                main.appConfig.favVis = checkBoxFav.Checked;

                // Fane: AVANSERT ---------------------------------------------------------------------------
                // AVANSERT - hoverdprodukt varegrupper
                if (VerifyStringToIntArray(textBoxHovedMda.Text))
                    main.appConfig.mainproductListMda = StringToIntArray(textBoxHovedMda.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxHovedAv.Text))
                    main.appConfig.mainproductListAv = StringToIntArray(textBoxHovedAv.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxHovedSda.Text))
                    main.appConfig.mainproductListSda = StringToIntArray(textBoxHovedSda.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxHovedTele.Text))
                    main.appConfig.mainproductListTele = StringToIntArray(textBoxHovedTele.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxHovedData.Text))
                    main.appConfig.mainproductListData = StringToIntArray(textBoxHovedData.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // AVANSERT - tilbehør hovedprodukter
                if (VerifyStringToIntArray(textBoxAccMda.Text))
                    main.appConfig.accessoriesListMda = StringToIntArray(textBoxAccMda.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxAccAv.Text))
                    main.appConfig.accessoriesListAv = StringToIntArray(textBoxAccAv.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxAccSda.Text))
                    main.appConfig.accessoriesListSda = StringToIntArray(textBoxAccSda.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxAccTele.Text))
                    main.appConfig.accessoriesListTele = StringToIntArray(textBoxAccTele.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (VerifyStringToIntArray(textBoxAccData.Text))
                    main.appConfig.accessoriesListData = StringToIntArray(textBoxAccData.Text);
                else
                {
                    tabControlMain.SelectedTab = tabPageAvansert;
                    Logg.Alert("Varegruppe skal være på tre siffer og ikke inneholde bokstaver.\nFlere varegrupper deles med komma.", "Feil varegruppe format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Fane: E-POST
                // E-POST - Server
                main.appConfig.epostSMTPserver = textBoxEpostSMTPserver.Text;
                main.appConfig.epostSMTPport = Convert.ToInt32(textBoxEpostSMTPport.Text);
                main.appConfig.epostSMTPssl = checkBoxEmailUseSsl.Checked;
                main.appConfig.epostBrukBcc = checkBoxEmailUseBcc.Checked;
                if (!IsValidEmail(textBoxEpostAvsender.Text) && textBoxEpostAvsender.Text != "")
                {
                    tabControlMain.SelectedTab = tabPageEpost;
                    textBoxEpostAvsender.ForeColor = Color.Red;
                    Logg.Alert("E-post avsender er ikke i gyldig format.\nEksempel: noreply@elkjop.no", "Ugyldig E-post adresse", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                else
                {
                    textBoxEpostAvsender.ForeColor = Color.Black;
                    main.appConfig.epostAvsender = textBoxEpostAvsender.Text;
                }
                main.appConfig.epostAvsenderNavn = textBoxEpostAvsenderNavn.Text;
                main.appConfig.epostPOP3server = textBoxEpostPOP3server.Text;
                main.appConfig.epostPOP3port = Convert.ToInt32(textBoxEpostPOP3port.Text);
                if (main.appConfig.epostPOP3port == 0)
                    main.appConfig.epostPOP3port = 995;
                main.appConfig.epostPOP3ssl = checkBoxPOP3UseSsl.Checked;
                main.appConfig.epostPOP3username = textBoxEpostPOP3bruker.Text;
                var aes = new SimpleAES();
                if (textBoxEpostPOP3pass.Text != "")
                    main.appConfig.epostPOP3password = aes.EncryptToString(textBoxEpostPOP3pass.Text);
                else
                    main.appConfig.epostPOP3password = "";
                main.appConfig.epostPOP3searchLimit = Convert.ToInt32(numericEpostPOP3searchLimit.Value);

                // E-POST - Annet
                main.appConfig.epostNesteMelding = textBoxEpostNesteMelding.Text.Replace(Environment.NewLine, "\n");

                // Fane: SERVICE -------------------------------------------------------------------------
                // SERVICE - Service generelt
                main.appConfig.serviceEgenServiceFilter = textServiceEgenServiceFilter.Text;
                main.appConfig.serviceShowHistory = checkServiceShowHistory.Checked;
                main.appConfig.serviceShowHistoryGraph = checkServiceShowHistoryGraph.Checked;
                main.appConfig.serviceShowTrend = checkServiceShowTrend.Checked;
                main.appConfig.serviceHistoryDays = Convert.ToInt32(numericServiceHistoryDays.Value);
                main.appConfig.serviceFerdigServiceStats = checkBoxServiceFerdigService.Checked;
                main.appConfig.serviceFerdigServiceStatsAntall = (int)numericServiceFerdigStatsAntall.Value;


                // Fane: GRAF ---------------------------------------------------------------------------
                // GRAF - Grafikk
                main.appConfig.graphVis = checkBoxVisGrafikk.Checked;
                main.appConfig.graphScreenDPI = (float)numericGraphDPI.Value;
                main.appConfig.graphWidth = Convert.ToInt32(textBoxGraphWidth.Text);
                main.appConfig.graphExtra = checkBoxGrafikkExtra.Checked;
                main.appConfig.graphAdvanced = checkBoxGrafikkAdvanced.Checked;
                main.appConfig.graphKnowHow = checkBoxGrafikkKnowHow.Checked;
                main.appConfig.graphButikk = checkBoxGrafikkButikk.Checked;
                main.appConfig.graphData = checkBoxGrafikkData.Checked;
                main.appConfig.graphAudioVideo = checkBoxGrafikkAV.Checked;
                main.appConfig.graphTele = checkBoxGrafikkTele.Checked;
                main.appConfig.graphOversikt = checkBoxGrafikkOversikt.Checked;
                try
                {
                    var tmp = Convert.ToInt32(textBoxGraphAntall.Text);
                    if (tmp > 0 && tmp < 366)
                        main.appConfig.graphDager = tmp;
                    if (Convert.ToInt32(textBoxGraphX.Text) > 50 && Convert.ToInt32(textBoxGraphY.Text) > 50)
                    {
                        main.appConfig.graphResX = Convert.ToInt32(textBoxGraphX.Text);
                        main.appConfig.graphResY = Convert.ToInt32(textBoxGraphY.Text);
                    }
                    else
                    {
                        main.appConfig.graphResX = 2000;
                        main.appConfig.graphResY = 600;
                    }
                }
                catch
                {
                    main.appConfig.graphResX = 2000;
                    main.appConfig.graphResY = 600;
                    main.appConfig.graphDager = 40;
                }

                // Fane: LAGER ------------------------------------------------------------------------------------
                // LAGER - Generelt
                main.appConfig.storeCompareMtd = radioLagerUtviklingMTD.Checked;
                main.appConfig.dbStoreViewpoint = pickerLagerViewpoint.Value;
                main.appConfig.storeShowStoreTwo = checkBoxStoreShowStoreTwo.Checked;
                main.appConfig.storeMaxAgePrizes = (int)numericUpDownLagerPriserMaxDager.Value;

                // LAGER - Utgåtte varer
                main.appConfig.storeObsoleteListMainProductsOnly = checkLagerHovedprodukter.Checked;
                main.appConfig.storeObsoleteFilterPercent = (int)numericStoreObsoleteFilter.Value;
                main.appConfig.storeObsoleteFilterMainStoreOnly = checkBoxObsoleteFilterMainStoreOnly.Checked;
                main.appConfig.storeObsoleteSortBy = comboBoxStoreObsoleteSortBy.Text;
                main.appConfig.storeObsoleteSortAsc = checkBoxStoreObsoleteSortAsc.Checked;
                main.appConfig.storeObsoleteFilterMax = (int)numericStoreObsoleteFilterMax.Value;

                // LAGER - Prisguide.no og Ukenytt
                main.appConfig.onlinePrisguidePagesToImport = (int)numericUpDownLagerPrisguideImportPages.Value;

                // Fane: WEBSERVER --------------------------------------------------------------------------------
                // WEBSERVER - Generelt
                main.appConfig.webserverEnabled = checkBoxWebserverEnabled.Checked;
                main.appConfig.webserverPort = (int)numericWebserverPort.Value;
                main.appConfig.webserverHost = comboBoxWebserverBindings.Text;

                // WEBSERVER - Sikkerhet
                main.appConfig.webserverRequireSimpleAuthentication = checkBoxWebserverSimpleAut.Checked;
                main.appConfig.webserverUser = textBoxWebserverUser.Text;
                if (textBoxWebserverPassword.Text != "")
                    main.appConfig.webserverPassword = aes.EncryptToString(textBoxWebserverPassword.Text);
                else
                    main.appConfig.webserverPassword = "";

                // Fane: VEDLIKEHOLD -------------------------------------------------------------------------------
                // VEDLIKEHOLD - Annet
                main.appConfig.debug = debugCheckBox.Checked;

                main.SaveSettings();
                SendMessage("Innstillinger lagret.", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                SendMessage(ex.Message.ToString().Replace("\n", "").Replace("\r", ""), Color.Red);
                Logg.Unhandled(ex);
                return false;
            }
        }

        private void DBsize()
        {
            try
            {
                labelDatabaseVersion.Text = FormMain.databaseVersion;

                long len = new FileInfo(FormMain.fileDatabase).Length;
                labelSizeMainDb.Text = BytesToString(len);
            }
            catch (Exception ex)
            {
                labelSizeMainDb.Text = "[Tom]";
                Logg.Unhandled(ex);
            }
        }

        public void SendMessage(string str, Color? c = null)
        {
            try
            {
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

                if (str != "")
                    Logg.Log(str, c, true);
            }
            catch
            {
                errorMessage.ForeColor = Color.Red;
                errorMessage.Text = "Feil i meldingssystem!" + Environment.NewLine + "Siste feilmelding: " + str;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBoxFavLeggtil.Text.Length == 4 && main.appConfig.Avdeling.ToString() != textBoxFavLeggtil.Text)
                {
                    
                    for (int i = 0; i < listBoxFavAvd.Items.Count; i++)
                    {
                        if (listBoxFavAvd.Items[i].ToString().StartsWith(textBoxFavLeggtil.Text))
                        {
                            listBoxFavAvd.Items.RemoveAt(i);
                            break;
                        }
                    }
                    Avdeling avdObject = new Avdeling();
                    listBoxFavAvd.Items.Add(textBoxFavLeggtil.Text + ": " + avdObject.Get(textBoxFavLeggtil.Text));
                    textBoxFavLeggtil.Text = "";
                }
            }
            catch
            {
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listBoxFavAvd.SelectedItems.Count;i++ )
                listBoxFavAvd.Items.Remove(listBoxFavAvd.SelectedItems[i]);
        }

        private void UpdateColor(object sender, EventArgs e)
        {
            try
            {
                var control = (Button)sender;
                colorDialog1.Color = control.BackColor;
                if (colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    string text = control.Text;
                    control.BackColor = colorDialog1.Color;
                    Color c = colorDialog1.Color;
                    var l = 0.333 * c.R + 0.333 * c.G + 0.333 * c.B;
                    if (l > 150)
                    {
                        control.ForeColor = Color.Black;
                        if (text == "color1")
                        {
                            main.appConfig.color1 = c;
                            main.appConfig.color1inv = false;
                        }
                        if (text == "color2")
                        {
                            main.appConfig.color2 = c;
                            main.appConfig.color2inv = false;
                        }
                        if (text == "color3")
                        {
                            main.appConfig.color3 = c;
                            main.appConfig.color3inv = false;
                        }
                        if (text == "color4")
                        {
                            main.appConfig.color4 = c;
                            main.appConfig.color4inv = false;
                        }
                        if (text == "color5")
                        {
                            main.appConfig.color5 = c;
                            main.appConfig.color5inv = false;
                        }
                        if (text == "color6")
                        {
                            main.appConfig.color6 = c;
                            main.appConfig.color6inv = false;
                        }
                        if (text == "ncolor1")
                        {
                            main.appConfig.ncolor1 = c;
                            main.appConfig.ncolor1inv = false;
                        }
                        if (text == "ncolor2")
                        {
                            main.appConfig.ncolor2 = c;
                            main.appConfig.ncolor2inv = false;
                        }
                        if (text == "ncolor3")
                        {
                            main.appConfig.ncolor3 = c;
                            main.appConfig.ncolor3inv = false;
                        }
                        if (text == "ncolor4")
                        {
                            main.appConfig.ncolor4 = c;
                            main.appConfig.ncolor4inv = false;
                        }
                        if (text == "ncolor5")
                        {
                            main.appConfig.ncolor5 = c;
                            main.appConfig.ncolor5inv = false;
                        }
                        if (text == "ncolor6")
                        {
                            main.appConfig.ncolor6 = c;
                            main.appConfig.ncolor6inv = false;
                        }
                    }
                    else
                    {
                        control.ForeColor = Color.White;
                        if (text == "color1")
                        {
                            main.appConfig.color1 = c;
                            main.appConfig.color1inv = true;
                        }
                        if (text == "color2")
                        {
                            main.appConfig.color2 = c;
                            main.appConfig.color2inv = true;
                        }
                        if (text == "color3")
                        {
                            main.appConfig.color3 = c;
                            main.appConfig.color3inv = true;
                        }
                        if (text == "color4")
                        {
                            main.appConfig.color4 = c;
                            main.appConfig.color4inv = true;
                        }
                        if (text == "color5")
                        {
                            main.appConfig.color5 = c;
                            main.appConfig.color5inv = true;
                        }
                        if (text == "color6")
                        {
                            main.appConfig.color6 = c;
                            main.appConfig.color6inv = true;
                        }
                        if (text == "ncolor1")
                        {
                            main.appConfig.ncolor1 = c;
                            main.appConfig.ncolor1inv = true;
                        }
                        if (text == "ncolor2")
                        {
                            main.appConfig.ncolor2 = c;
                            main.appConfig.ncolor2inv = true;
                        }
                        if (text == "ncolor3")
                        {
                            main.appConfig.ncolor3 = c;
                            main.appConfig.ncolor3inv = true;
                        }
                        if (text == "ncolor4")
                        {
                            main.appConfig.ncolor4 = c;
                            main.appConfig.ncolor4inv = true;
                        }
                        if (text == "ncolor5")
                        {
                            main.appConfig.ncolor5 = c;
                            main.appConfig.ncolor5inv = true;
                        }
                        if (text == "ncolor6")
                        {
                            main.appConfig.ncolor6 = c;
                            main.appConfig.ncolor6inv = true;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Logg.Debug("Noe uventet skjedde i UpdateColor()", ex);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            ResetSettings();
        }

        private void ResetSettings()
        {
            try
            {
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle innstillinger?", "KGSA - ADVARSEL!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    main.appConfig = new AppSettings();
                    main.SaveSettings();
                    ImportAppSettings();
                    textBoxFavLeggtil.Text = "";
                    tabControlMain.SelectedTab = tabPageGenerelt;
                    SendMessage("Innstillinger tilbakestilt.", Color.Green);
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void textBoxFavLeggtil_Enter(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = button14;
        }

        private void textBoxFavLeggtil_Leave(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = null;
        }

        private void textBoxAvd_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (textBoxAvd.TextLength == 4)
                {
                    int temp = Convert.ToInt32(textBoxAvd.Text);
                    if (temp >= 1000 && temp <= 1700)
                    {
                        label19.Text = main.avdeling.Get(temp);
                        main.appConfig.Avdeling = temp;
                        button1.Enabled = true;
                        textBoxAvd.BackColor = ColorTranslator.FromHtml("#b9dc9c");
                        errorMessage.Text = "";
                    }
                }
                else
                {
                    button1.Enabled = false;
                    label19.Text = "";
                    textBoxAvd.BackColor = ColorTranslator.FromHtml("#ecc3ad");
                }
            }
            catch
            {
                SendMessage("Ugyldig avdelingsnummer.", Color.Red);
            }
        }

        private void buttonCSVinegoBrowse_Click(object sender, EventArgs e)
        {
            // Browse etter elguide eksport mappe
            try
            {

                var fbd = new FolderBrowserDialog();
                fbd.Description = "Velg mappen som Elguide eksporterer som standard, vanligvis C:/";
                fbd.ShowNewFolderButton = false;
                fbd.RootFolder = Environment.SpecialFolder.MyComputer;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    textBoxCSVexportFolder.Text = fbd.SelectedPath;
                }
                fbd.Dispose();
            }
            catch(Exception ex)
            {
                SendMessage("Unntak ved Elguide eksport mappe valg.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            if (textBoxAvd.Text.Length != 4)
                this.ActiveControl = textBoxAvd;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            EraseSelgerkoder();
        }

        private void EraseSelgerkoder()
        {
            try
            {
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle selgerkoder fra databasen?",
                    "KGSA - ADVARSEL",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    SendMessage("Vent litt..");

                    main.database.tableSelgerkoder.Reset();
                    main.sKoder = new Selgerkoder(main, true);

                    SendMessage("Database: Alle selgerkoder slettet.");
                    Logg.Log("Database: Selgerkode tabelllen nullstilt.");
                    DBsize();
                }
            }
            catch (Exception ex)
            {
                SendMessage("Unntak oppstod ved sletting av varekoder.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EraseRanking();
        }

        private void EraseRanking()
        {
            try
            {
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle transaksjoner fra databasen?",
                    "KGSA - VIKTIG",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {

                    SendMessage("Vent litt..");
                    Logg.Log("Database: Sletter salgs tabellen..");

                    main.database.tableSalg.Reset();

                    main.appConfig.dbFrom = DateTime.MinValue;
                    main.appConfig.dbTo = DateTime.MinValue;
                    main.SaveSettings();
                    DBsize();
                    SendMessage("Transaksjons tabellen slettet.");
                    Logg.Log("Database: Operasjon utført.");
                    forceUpdate = true;
                }
            }
            catch (Exception ex)
            {
                SendMessage("Unntak oppstod ved sletting av transaksjoner.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void button8_Click_2(object sender, EventArgs e)
        {
            EraseService();
        }

        private void EraseService()
        {
            try
            {
                if (main.service.ClearDatabase())
                {
                    SendMessage("Service databasen tømt.", Color.Green);
                    forceUpdate = true;
                }
            }
            catch (Exception ex)
            {
                SendMessage("Unntak oppstod ved sletting av service databasen.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("explorer.exe", FormMain.settingsPath);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                System.IO.DirectoryInfo downloadedMessageInfo = new DirectoryInfo(System.IO.Path.GetTempPath());
                int teller = 0;
                foreach (FileInfo file in downloadedMessageInfo.GetFiles())
                {
                    if (file.Name.Contains("KGSA") && file.Name.Contains(".csv"))
                    {
                        file.Delete();
                        teller++;
                    }
                }
                SendMessage("Slettet " + teller + " csv.", Color.Green);
                Logg.Log("Tømt mellomlager for " + teller + " KGSA CSV filer.");
            }
            catch(Exception ex)
            {
                Logg.Debug("Unntak under sletting av CSV.", ex);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            CompactDb();
        }

        private void CompactDb()
        {
            try
            {
                main.database.CloseConnection();
                SendMessage("Compacting DB.. vent litt" + Environment.NewLine + "(Dette kan ta en stund)");
                this.Update();
                Logg.Log("Compacting DB.. vent litt");

                SqlCeEngine eng = new SqlCeEngine(FormMain.SqlConStr);
                eng.Compact(FormMain.SqlConStr);
                //eng.Compact("Data Source=" + Path.Combine(FormMain.settingsPath, @"\db_compacting.sdf"));

                main.database.OpenConnection();

                DBsize();

                Logg.Log("Compacting av " + FormMain.fileDatabase + " fullført.", Color.Green);
                SendMessage("Ferdig.", Color.Green);

            }
            catch (Exception ex)
            {
                SendMessage("Feil!", Color.Red);
                FormError errorMsg = new FormError("Feil oppstod under database optimalisering", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void comboBoxImport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxImport.SelectedIndex == 0)
                labelImportForklaring.Text = "Importere bare transaksjoner som er relevant for\nsalg av klargjøringer og supportavtaler.\nEr det raskeste alternativet.";
            if (comboBoxImport.SelectedIndex == 1)
                labelImportForklaring.Text = "Importere alle transaksjoner, også fra alle andre\navdelinger hvis de finnes. Er nødvendig\nfor ranking av andre tjenester.";
            if (comboBoxImport.SelectedIndex == 2)
                labelImportForklaring.Text = "Importere alle transaksjoner fra din egen avdeling\nsamt angitte favoritt avdelinger. \nEr nødvendig for ranking av andre tjenester.";
            if (main.appConfig.importSetting == "Normal" && !FormMain.EmptyDatabase() && comboBoxImport.SelectedIndex > 0)
                MessageBox.Show("Viktig: Hvis du har importert tidligere uten å ta med alle transaksjonene vil ranking på de andre tjeneste utenom KG/SA være ufullstendige. Importer transaksjonene for hele perioden på nytt.", "KGSA - Informasjon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonTestUnlock_Click(object sender, EventArgs e)
        {

            bwUnlockTest.RunWorkerAsync();
        }

        private void button21_Click_1(object sender, EventArgs e)
        {
            // Browse etter pdf eksport mappe
            try
            {
                var fbd = new FolderBrowserDialog();
                fbd.Description = "Velg mappen som KGSA skal lagre PDF rankinger";
                fbd.ShowNewFolderButton = false;
                fbd.RootFolder = Environment.SpecialFolder.MyComputer;
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    textBoxExportPDF.Text = fbd.SelectedPath;
                }
                fbd.Dispose();
            }
            catch
            {
                SendMessage("Unntak ved PDF eksport mappe valg.", Color.Red);
            }
        }

        private void numericM_ValueChanged(object sender, EventArgs e)
        {
            var control = (NumericUpDown)sender;
            if (control.Value == 0)
                control.ForeColor = Color.Gray;
            else
                control.ForeColor = Color.Black;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            EraseStore();
        }

        private void EraseStore()
        {
            try
            {
                if (main.obsolete.ClearDatabase())
                {
                    SendMessage("Lager databasen nullstilt.", Color.Green);
                    forceUpdate = true;
                }
            }
            catch (Exception ex)
            {
                SendMessage("Unntak oppstod ved sletting av lager databasen.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void buttonHentButikker_Click(object sender, EventArgs e)
        {
            PopulateAlleButikkerListe();
        }

        private void PopulateAlleButikkerListe()
        {
            try
            {
                if (FormMain.EmptyDatabase())
                {
                    SendMessage("Databasen er tom!", Color.Red);
                    return;
                }

                SendMessage("Henter butikker..");
                listBoxButikkerDatabase.Items.Clear();
                listBoxButikkerDatabase.Items.Add("Laster..");
                this.Update();
                DataTable dt = main.database.GetSqlDataTable("SELECT DISTINCT Avdeling FROM tblSalg WHERE Avdeling < 1700 AND Dato >= '" + FormMain.dbTilDT.AddYears(-1).ToString("yyy-MM-dd") + "' AND Dato <= '" + FormMain.dbTilDT.ToString("yyy-MM-dd") + "'");
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        listBoxButikkerDatabase.Items.Clear();
                        Avdeling avdObject = new Avdeling();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string str = avdObject.Get(dt.Rows[i][0].ToString());
                            if (str.Length == 4)
                                listBoxButikkerDatabase.Items.Add(dt.Rows[i][0]);
                            else
                                listBoxButikkerDatabase.Items.Add(dt.Rows[i][0] + ": " + str);
                        }
                        SendMessage("Butikkliste oppdatert.", Color.Green);
                    }
                    else
                    {
                        listBoxButikkerDatabase.Items.Clear();
                        SendMessage("Ingen butikker funnet!", Color.Red);
                    }
                }
                else
                    listBoxButikkerDatabase.Items.Clear();
            }
            catch(Exception ex)
            {
                SendMessage("Klarte ikke å oppdatere butikk listen.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void buttonButikkerLeggTilAlle_Click(object sender, EventArgs e)
        {
            FlyttButikkListeTilFavoritter(true);
        }

        private void FlyttButikkListeTilFavoritter(bool alle)
        {
            try
            {
                if (FormMain.EmptyDatabase() || listBoxButikkerDatabase.Items.Count == 0)
                {
                    SendMessage("Databasen er tom eller butikklisten er ikke oppdatert.", Color.Red);
                    return;
                }

                List<string> samleListe = new List<string> { };

                for (int b = 0; b < listBoxFavAvd.Items.Count; b++)
                    samleListe.Add(listBoxFavAvd.Items[b].ToString().Substring(0, 4));

                if (listBoxButikkerDatabase.Items.Count > 0)
                {
                    if (alle)
                    {
                        for (int i = 0; i < listBoxButikkerDatabase.Items.Count; i++)
                            samleListe.Add(listBoxButikkerDatabase.Items[i].ToString().Substring(0, 4));
                    }
                    else
                        for (int i = 0; i < listBoxButikkerDatabase.SelectedItems.Count; i++ )
                            samleListe.Add(listBoxButikkerDatabase.SelectedItems[i].ToString().Substring(0, 4));
                }

                if (samleListe.Count > 0)
                {
                    listBoxFavAvd.Items.Clear();
                    Avdeling avdObject = new Avdeling();
                    samleListe = samleListe.Distinct().ToList();
                    foreach (string avd in samleListe)
                    {
                        if (avd == main.appConfig.Avdeling.ToString())
                            continue;
                        string str = avdObject.Get(avd);
                        if (str.Length == 4)
                            listBoxFavAvd.Items.Add(avd);
                        else
                            listBoxFavAvd.Items.Add(avd + ": " + str);
                    }
                    listBoxButikkerDatabase.ClearSelected();
                    SendMessage("Butikker flyttet.", Color.Green);
                }
                else
                    SendMessage("Ingen butikker å flytte!?", Color.Red);
            }
            catch(Exception ex)
            {
                SendMessage("Klarte ikke å flytte over butikklisten!", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void listBoxButikkerDatabase_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (listBoxButikkerDatabase.SelectedIndex > -1)
                    if (listBoxButikkerDatabase.SelectedItem.ToString().Length >= 4)
                    {
                        string str = listBoxButikkerDatabase.SelectedItem.ToString().Substring(0, 4);
                        if (main.appConfig.Avdeling.ToString() != str)
                        {

                            for (int i = 0; i < listBoxFavAvd.Items.Count; i++)
                            {
                                if (listBoxFavAvd.Items[i].ToString().StartsWith(str))
                                {
                                    listBoxFavAvd.Items.RemoveAt(i);
                                    break;
                                }
                            }
                            Avdeling avdObject = new Avdeling();
                            listBoxFavAvd.Items.Add(str + ": " + avdObject.Get(str));
                        }
                    }
            }
            catch
            {
            }
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
                Logg.Unhandled(ex);
            }
        }

        private List<string> GetAllBindings()
        {
            try
            {
                IPHostEntry host;
                List<string> localBindings = new List<string>() { };
                host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (IPAddress ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        localBindings.Add(ip.ToString());
                    }
                }
                localBindings.Add("localhost");
                return localBindings;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                SendMessage("Error: Klarte ikke å hente ip-addresser.");
                return null;
            }
        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            var form = new FormEmailAddressbook(main);
            form.ShowDialog(this);
        }

        private void buttonValidateDb_Click(object sender, EventArgs e)
        {
            main.database.VerifyDatabase();
            SendMessage("Valdering av databasen fullført.", Color.Green);
        }

        private void checkBoxAdvanced_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAdvanced.Checked)
                AddAdvancedTabs();
            else
                RemoveAdvancedTabs();
        }

        private void AddAdvancedTabs()
        {
            if (!tabControlMain.TabPages.Contains(tabPageVedlikehold))
            {
                tabControlMain.TabPages.Add(tabPageLager);
                tabControlMain.TabPages.Add(tabPageService);
                tabControlMain.TabPages.Add(tabPageWebserver);
            }
        }

        private void RemoveAdvancedTabs()
        {
            if (tabControlMain.TabPages.Contains(tabPageVedlikehold))
            {
                tabControlMain.TabPages.Remove(tabPageLager);
                tabControlMain.TabPages.Remove(tabPageService);
                tabControlMain.TabPages.Remove(tabPageWebserver);
            }
        }

        private void AddExperimentalTabs()
        {
            if (!tabControlMain.TabPages.Contains(tabPageWebserver))
            {
                tabControlMain.TabPages.Add(tabPageWebserver);
            }
        }

        private void RemoveExperimentalTabs()
        {
            if (tabControlMain.TabPages.Contains(tabPageWebserver))
            {
                tabControlMain.TabPages.Remove(tabPageWebserver);
            }
        }

        private void buttonButikkerLeggTil_Click(object sender, EventArgs e)
        {
            FlyttButikkListeTilFavoritter(false);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!ExportSettings())
                SendMessage("Feil format/innstilling.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult != System.Windows.Forms.DialogResult.Cancel)
            {
                if (!ExportSettings() && this.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    SendMessage("Feil format/innstilling.");
                    e.Cancel = true;
                }
            }
        }

        /// <summary>
        /// Gjenoppbygging av tblVinnprodukt
        /// </summary>
        private void CleanVinnDb()
        {
            try
            {

                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle vinnprodukter fra databasen?",
                    "KGSA - VIKTIG",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    SendMessage("Vent litt..");
                    main.database.tableVinnprodukt.Reset();
                    SendMessage("Database: Operasjon fullført. Slettet alle vinnprodukter.", Color.Green);
                }
            }
            catch (Exception ex)
            {
                SendMessage("Feil oppstod under gjenoppretting av tabell.", Color.Red);
                Logg.Unhandled(ex);
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            CleanVinnDb();
        }

        private void checkBoxActivateExperimental_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxActivateExperimental.Checked)
                AddExperimentalTabs();
            else
                RemoveExperimentalTabs();
        }

        private int[] StringToIntArray(string input)
        {
            try
            {
                int[] output = input.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
                return output;
            }
            catch
            {
                SendMessage("Feil ved lagring av varegruppe.", Color.Red);
                return new int[] {0};
            }
        }

        private bool VerifyStringToIntArray(string input)
        {
            try
            {
                int[] output = input.Split(',').Select(n => Convert.ToInt32(n)).ToArray();
                return true;
            }
            catch
            {
                SendMessage("Feil format på varegruppe: " + input, Color.Red);
                return false;
            }
        }

        private string IntArrayToString(int[] input)
        {
            try
            {
                return String.Join(",", new List<int>(input).ConvertAll(i => i.ToString()).ToArray()); ;
            }
            catch
            {
                SendMessage("Feil ved henting av varegruppe.", Color.Red);
                return "";
            }
        }

        private bool VerifyIntArrayToString(int[] input)
        {
            try
            {
                string test = String.Join(",", new List<int>(input).ConvertAll(i => i.ToString()).ToArray());
                return true;
            }
            catch
            {
                SendMessage("Feil format på varegruppe: " + input, Color.Red);
                return false;
            }
        }

        private void buttonHovedReset_Click(object sender, EventArgs e)
        {
            ResetVaregruppeFelt(false);
        }

        private void ResetVaregruppeFelt(bool acc)
        {
            try
            {
                if (!acc)
                {
                    textBoxHovedMda.Text = IntArrayToString(new int[8] { 131, 132, 134, 135, 136, 143, 144, 145 });
                    textBoxHovedAv.Text = IntArrayToString(new int[2] { 224, 273 });
                    textBoxHovedSda.Text = IntArrayToString(new int[2] { 301, 346 });
                    textBoxHovedTele.Text = IntArrayToString(new int[2] { 431, 447 });
                    textBoxHovedData.Text = IntArrayToString(new int[3] { 531, 533, 534 });
                }
                else
                {
                    textBoxAccMda.Text = IntArrayToString(new int[1] { 195 });
                    textBoxAccAv.Text = IntArrayToString(new int[1] { 214 });
                    textBoxAccSda.Text = IntArrayToString(new int[1] { 395 });
                    textBoxAccTele.Text = IntArrayToString(new int[1] { 487 });
                    textBoxAccData.Text = IntArrayToString(new int[3] { 552, 569, 589 });
                }

                SendMessage("Varegruppe satt tilbake til standardverdier.");
            }
            catch
            {
                SendMessage("OBS! Unntak ved varegruppe reset", Color.Red);
                return;
            }
        }

        private void buttonAccReset_Click(object sender, EventArgs e)
        {
            ResetVaregruppeFelt(true);
        }

        private void pictureElkjop_Click_1(object sender, EventArgs e)
        {
            if (!main.appConfig.chainElkjop)
            {
                main.appConfig.chainElkjop = true;
                panelPictureElkjop.BackColor = Color.Black;
                panelPictureLefdal.BackColor = Color.White;
                SendMessage("Butikk kjede endret til Elkjøp.", Color.Green);
            }
            else
                SendMessage("");
        }

        private void pictureLefdal_Click_1(object sender, EventArgs e)
        {
            if (main.appConfig.chainElkjop)
            {
                main.appConfig.chainElkjop = false;
                panelPictureElkjop.BackColor = Color.White;
                panelPictureLefdal.BackColor = Color.Black;
                SendMessage("Butikk kjede endret til Lefdal.", Color.Green);
            }
            else
                SendMessage("");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            { 
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle EAN koder fra databasen?",
                    "KGSA - VIKTIG",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    SendMessage("Vent litt..");
                    main.database.tableEan.Reset();
                    SendMessage("Database: Operasjon fullført. EAN koder slettet.", Color.Green);
                }
            }
            catch (Exception ex)
            {
                SendMessage("Feil oppstod under gjenoppretting av tabell.", Color.Red);
                Logg.Unhandled(ex);
            }
        }

        private void buttonMaintenancePrintStatus_Click(object sender, EventArgs e)
        {
            main.database.PrintStatus();
        }
    }
}