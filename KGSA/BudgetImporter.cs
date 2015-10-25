using System;
using System.Collections.Generic;
using System.Linq;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using Message = OpenPop.Mime.Message;
using iTextSharp.text.pdf;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data;

namespace KGSA
{
    public class BudgetImporter
    {
        FormMain main;
        BackgroundWorker worker;
        private string matchSubjectMustContain = "c810";
        private string matchAttachmentFilenameMustContain = "c810";
        private int limitEmailAttempts = 5;
        private int countEmailAttempts = 0;
        private bool maxEmailAttemptsReached = false;
        private bool runInBackground = true;

        private DateTime selectedDate;

        public BudgetImporter(FormMain form, DateTime date)
        {
            selectedDate = date;
            main = form;
        }

        public void StartAsyncDownloadBudget(BackgroundWorker bw, bool runInBackground = true)
        {
            worker = bw;
            worker.DoWork += new DoWorkEventHandler(bwDownload_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressReport_ProgressChanged);
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwDownload_Completed);

            maxEmailAttemptsReached = false;
            countEmailAttempts = 0;

            main.processing.SetVisible = true;
            main.processing.SetText = "Forbereder..";
            main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
            this.runInBackground = runInBackground;
            worker.RunWorkerAsync();
            main.processing.SetBackgroundWorker = worker;
        }

        public void bwDownload_DoWork(object sender, DoWorkEventArgs e)
        {
            if (FindAndDownloadBudget(worker))
            {
                e.Result = true;
                Log.n("Dagens budsjett for " + selectedDate.ToShortDateString() + " er lastet ned og lagret", Color.Green);
                return;
            }
            e.Result = false;
        }

        public void bwDownload_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            main.ProgressStop();
            if (!e.Cancelled && e.Error == null && (bool)e.Result)
            {
                if (runInBackground)
                    main.RunBudget(BudgetCategory.Daglig);
                main.processing.SetVisible = false;
            }
            else if (!e.Cancelled && e.Error == null && !(bool)e.Result)
            {
                main.processing.SetVisible = false;
                if (runInBackground)
                    Log.n("Fant ikke dagens budsjett i innboks " + main.appConfig.epostPOP3username + ". Sjekk logg for detaljer.", Color.Red);
                else
                    Log.Alert("Fant ikke dagens budsjett i innboks " + main.appConfig.epostPOP3username + ".\nSjekk logg for detaljer.", "Nedlasting av Dagens budsjett", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (e.Cancelled || (worker != null && worker.CancellationPending))
            {
                Log.n("Prosessen ble stopped av bruker", Color.Red);
                main.processing.SetText = "Avbrutt";
                main.processing.HideDelayed();
            }
            else
                main.processing.SetVisible = false;

            this.runInBackground = true;
        }


        public DataTable ImportElguideBudget(int avdeling)
        {
            try
            {
                string path = Path.Combine(main.appConfig.csvElguideExportFolder, "inego.csv");
                if (!File.Exists(path))
                {
                    Log.n("Fant ikke CSV: " + path, Color.Red);
                    return null;
                }

                else if (File.GetLastWriteTime(path).Date != DateTime.Now.Date)
                {
                    Log.n("CSV var ikke oppdatert i dag. Eksporter fra Elguide program 136 først", Color.Red);
                    return null;
                }

                string[] Lines = File.ReadAllLines(path);
                string[] Fields;
                Fields = Lines[0].Split(new char[] { ';' });
                int Cols = Fields.GetLength(0);
                DataTable table = new DataTable();
                //1st row must be column names; force lower case to ensure matching later on.
                for (int i = 0; i < Cols; i++)
                    table.Columns.Add(Fields[i].ToLower(), typeof(string));
                DataRow Row;
                for (int i = 1; i < Lines.GetLength(0); i++)
                {
                    Fields = Lines[i].Split(new char[] { ';' });
                    Row = table.NewRow();
                    for (int f = 0; f < Cols; f++)
                        Row[f] = Fields[f];
                    table.Rows.Add(Row);
                }

                if (table.Rows.Count < 6)
                {
                    Log.n("CSV " + Path.GetFileName(path) + " inneholder fem eller mindre linjer. Eksporter fra Elguide på nytt og prøv igjen", Color.Red);
                    return null;
                }

                DataTable tableQuick = new DataTable();
                tableQuick.Columns.Add("Favoritt", typeof(int));
                tableQuick.Columns.Add("Avdeling", typeof(string));
                tableQuick.Columns.Add("Salg", typeof(int));
                tableQuick.Columns.Add("Omsetn", typeof(decimal));
                tableQuick.Columns.Add("Fritt", typeof(decimal));
                tableQuick.Columns.Add("Fortjeneste", typeof(decimal));
                tableQuick.Columns.Add("Margin", typeof(double));
                tableQuick.Columns.Add("Rabatt", typeof(decimal));

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (table.Rows[i][0].ToString().Contains("---"))
                    {
                        // Vi er i row før totalen.
                        i++;
                        DataRow rowTot = tableQuick.NewRow();
                        rowTot["Favoritt"] = avdeling;
                        rowTot["Avdeling"] = "TOTAL";
                        rowTot["Salg"] = Convert.ToInt32(table.Rows[i][1].ToString());
                        rowTot["Omsetn"] = Convert.ToDecimal(table.Rows[i][2].ToString());
                        rowTot["Fritt"] = Convert.ToDecimal(table.Rows[i][3].ToString());
                        rowTot["Fortjeneste"] = Convert.ToDecimal(table.Rows[i][4].ToString());
                        rowTot["Margin"] = Convert.ToDouble(table.Rows[i][5].ToString());
                        rowTot["Rabatt"] = Convert.ToDecimal(table.Rows[i][6].ToString());
                        tableQuick.Rows.Add(rowTot);
                        break;
                    }

                    DataRow row = tableQuick.NewRow();
                    row["Favoritt"] = avdeling;
                    string tmp = table.Rows[i][0].ToString().Substring(1, 1);
                    if (tmp == "1")
                        tmp = "MDA";
                    if (tmp == "2")
                        tmp = "AudioVideo";
                    if (tmp == "3")
                        tmp = "SDA";
                    if (tmp == "4")
                        tmp = "Telecom";
                    if (tmp == "5")
                        tmp = "Computing";
                    if (tmp == "6")
                        tmp = "Kitchen";
                    if (tmp == "9")
                        tmp = "Other";
                    row["Avdeling"] = tmp;
                    if (!String.IsNullOrEmpty(table.Rows[i][1].ToString()))
                        row["Salg"] = Convert.ToInt32(table.Rows[i][1].ToString());
                    else
                        row["Salg"] = 0;
                    if (!String.IsNullOrEmpty(table.Rows[i][2].ToString()))
                        row["Omsetn"] = Convert.ToDecimal(table.Rows[i][2].ToString());
                    else
                        row["Omsetn"] = 0;
                    if (!String.IsNullOrEmpty(table.Rows[i][3].ToString()))
                        row["Fritt"] = Convert.ToDecimal(table.Rows[i][3].ToString());
                    else
                        row["Fritt"] = 0;
                    if (!String.IsNullOrEmpty(table.Rows[i][4].ToString()))
                        row["Fortjeneste"] = Convert.ToDecimal(table.Rows[i][4].ToString());
                    else
                        row["Fortjeneste"] = 0;
                    if (!String.IsNullOrEmpty(table.Rows[i][5].ToString()))
                        row["Margin"] = Convert.ToDouble(table.Rows[i][5].ToString());
                    else
                        row["Margin"] = 0;
                    if (!String.IsNullOrEmpty(table.Rows[i][6].ToString()))
                        row["Rabatt"] = Convert.ToDecimal(table.Rows[i][6].ToString());
                    else
                        row["Rabatt"] = 0;
                    tableQuick.Rows.Add(row);
                }

                return tableQuick;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Uventet feil under importering av CSV fra avdeling " + avdeling + ". Feilmelding: " + ex.Message, Color.Red);
            }
            return null;
        }

        public DailyBudgetMacroInfo ImportElguideServices()
        {
            try
            {
                DailyBudgetMacroInfo budgetInfo = new DailyBudgetMacroInfo();

                string[] listOfFiles = new string[] { "inegoAV.csv", "inegoTele.csv", "inegoComp.csv" };

                foreach (string file in listOfFiles)
                {
                    if (!File.Exists(main.appConfig.csvElguideExportFolder + file))
                    {
                        Log.n("Fant ikke " + file, Color.Red);
                        return null;
                    }
                    else if (File.GetLastWriteTime(main.appConfig.csvElguideExportFolder + file).Date != DateTime.Now.Date)
                    {
                        Log.n("Avbryter importering av " + file + " fordi filen er ikke oppdatert!", Color.Red);
                        return null;
                    }

                    string[] Lines = File.ReadAllLines(main.appConfig.csvElguideExportFolder + file);
                    string[] Fields;
                    Fields = Lines[0].Split(new char[] { ';' });
                    int Cols = Fields.GetLength(0);
                    DataTable dt = new DataTable();
                    //1st row must be column names; force lower case to ensure matching later on.
                    for (int i = 0; i < Cols; i++)
                        dt.Columns.Add(Fields[i].ToLower(), typeof(string));
                    DataRow Row;
                    for (int i = 1; i < Lines.GetLength(0); i++)
                    {
                        Fields = Lines[i].Split(new char[] { ';' });
                        Row = dt.NewRow();
                        for (int f = 0; f < Cols; f++)
                            Row[f] = Fields[f];
                        dt.Rows.Add(Row);
                    }

                    if (dt.Rows.Count < 2)
                    {
                        Log.n("Sales CSV fil (" + file + ") var for kort. " + dt.Rows.Count + " linjer.", Color.Red);
                        return null;
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        if (dt.Rows[i][0].ToString().Contains("580"))
                        {
                            DailyBudgetMacroInfoItem salg = new DailyBudgetMacroInfoItem();
                            salg.Type = "Computing";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            budgetInfo.Salg.Add(salg);
                            Log.d("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }

                        if (dt.Rows[i][0].ToString().Contains("480"))
                        {
                            DailyBudgetMacroInfoItem salg = new DailyBudgetMacroInfoItem();
                            salg.Type = "Telecom";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            budgetInfo.Salg.Add(salg);
                            Log.d("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }

                        if (dt.Rows[i][0].ToString().Contains("280"))
                        {
                            DailyBudgetMacroInfoItem salg = new DailyBudgetMacroInfoItem();
                            salg.Type = "AudioVideo";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            budgetInfo.Salg.Add(salg);
                            Log.d("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }
                    }
                }
                return budgetInfo;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil oppstod under importering av tjeneste CSV fra Elguide", Color.Red);
            }
            return null;
        }

        public bool FindAndDownloadBudget(BackgroundWorker bw)
        {
            bool attachmentFound = false;
            try
            {
                Log.n("Kobler til e-post server \"" + main.appConfig.epostPOP3server + "\"..");
                using (Pop3Client client = new Pop3Client())
                {
                    client.Connect(main.appConfig.epostPOP3server, main.appConfig.epostPOP3port, main.appConfig.epostPOP3ssl);

                    if (bw != null) {
                        if (bw.CancellationPending) {
                            if (client.Connected)
                                client.Disconnect();
                            return false;
                        }
                    }

                    Log.d("Logger inn med brukernavn \"" + main.appConfig.epostBrukernavn + "\"..");
                    SimpleAES aes = new SimpleAES();
                    client.Authenticate(main.appConfig.epostPOP3username, aes.DecryptString(main.appConfig.epostPOP3password));
                    int count = client.GetMessageCount();

                    Log.d("Innlogging fullført. Antall meldinger i innboks: " + count);
                    if (count == 0)
                    {
                        Log.n("Innboks for \"" + main.appConfig.epostPOP3username + "\" er tom. "
                            + "Endre innstillinger i ditt e-post program til å la e-post ligge igjen på server", Color.Red);
                        client.Disconnect();
                        return false;
                    }

                    main.processing.SetText = "Søker i innboks.. (totalt " + count + " meldinger)";
                    Log.n("Søker i innboks.. (totalt " + count + " meldinger)");
                    for (int i = count + 1; i-- > 1;)
                    {
                        if (attachmentFound || maxEmailAttemptsReached)
                            break;

                        if (bw != null)
                            if (bw.CancellationPending)
                                break;

                        Log.d("Sjekker meldingshode " + (count - i + 1) + "..");
                        MessageHeader header = client.GetMessageHeaders(i);
                        if (HeaderMatch(header))
                        {
                            if (header.DateSent.Date != selectedDate.Date)
                            {
                                Log.d("Fant C810 e-post med annen dato: " + header.DateSent.ToShortDateString()
                                    + " ser etter: " + selectedDate.ToShortDateString() + " Emne: \"" + header.Subject
                                    + "\" Fra: \"" + header.From.MailAddress + "\"");
                                continue;
                            }

                            Log.d("--------- Fant C810 e-post kandidat: Fra: \""
                                + header.From.MailAddress.Address + "\" Emne: \"" + header.Subject
                                + "\" Sendt: " + header.DateSent.ToShortDateString() + " (" + header.DateSent.ToShortTimeString() + ")");

                            Log.d("Laster ned e-post # " + i + "..");
                            Message message = client.GetMessage(i);

                            foreach (MessagePart attachment in message.FindAllAttachments())
                            {                                    
                                if (attachmentFound)
                                    break;

                                Log.d("Vedlegg: " + attachment.FileName);
                                if (AttachmentMatch(attachment))
                                    attachmentFound = ParseAndSaveDailyBudget(attachment,
                                        message.Headers.DateSent,
                                        message.Headers.Subject,
                                        message.Headers.From.MailAddress.Address);
                            }
                        }

                        if (!attachmentFound)
                            Log.d("Fant ingen C810 vedlegg i e-post # " + i);
                    }
                    client.Disconnect();
                }

                if (attachmentFound)
                    return true;
                else if (!attachmentFound && maxEmailAttemptsReached)
                    Log.n("Maksimum antall e-post nedlastninger er overgått. Innboks søk er avsluttet", Color.Red);
                else
                    Log.n("Fant ingen C810 e-post med budsjett for dato " + selectedDate.ToShortDateString() + ". Innboks søk er avsluttet", Color.Red);
            }
            catch (PopServerNotFoundException)
            {
                Log.n("E-post server finnes ikke eller DNS problem ved tilkobling til " + main.appConfig.epostPOP3server, Color.Red);
            }
            catch (PopServerNotAvailableException ex)
            {
                Log.n("E-post server \"" + main.appConfig.epostPOP3server + "\" svarer ikke. Feilmelding: " + ex.Message, Color.Red);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Authentication failed") || ex.Message.Contains("credentials"))
                    Log.n("Feil brukernavn eller passord angitt for " + main.appConfig.epostPOP3username + ". Sjekk e-post innstillinger!", Color.Red);
                else
                {
                    Log.Unhandled(ex);
                    Log.n("Feil ved henting av epost. Beskjed: " + ex.Message);
                }
            }
            return false;
        }

        private bool ParseAndSaveDailyBudget(MessagePart attachment, DateTime sentDate, string subject, string address)
        {
            KpiBudget kpiBudget = ParseC810(attachment.Body);
            if (kpiBudget != null)
            {
                Log.n("Vedlegget '" + attachment.FileName + "' fra e-posten: \""
                    + subject + "\" fra: \"" + address + "\" sendt: " + sentDate.ToShortDateString() + " er avlest og verifisert.");
                if (SaveBudget(kpiBudget))
                {
                    Log.d("Budsjett lagret til databasen");
                    return true;
                }
            }
            countEmailAttempts++;
            if (countEmailAttempts > limitEmailAttempts)
                maxEmailAttemptsReached = true;

            Log.n("Vedlegget '" + attachment.FileName + "' fra e-posten: \""
                + subject + "\" fra: \"" + address + "\" sendt: " + sentDate.ToShortDateString() + " er ugyldig, feil eller var ikke av typen C810");
            return false;
        }

        private bool SaveBudget(KpiBudget kpiBudget)
        {
            try
            {
                if (kpiBudget == null || kpiBudget.element == null || kpiBudget.element.Count == 0)
                {
                    Log.n("Budsjett data er ubrukelig og kan ikke lagres", Color.Red);
                    return false;
                }

                DataTable table = main.database.tableDailyBudget.GetDataTable();
                TimeSpan elapsed = kpiBudget.Date.Subtract(FormMain.rangeMin);
                int budgetId = (int)elapsed.TotalDays;

                foreach (KpiBudgetElement element in kpiBudget.element)
                {
                    DataRow dtRow = table.NewRow();

                    dtRow[TableDailyBudget.KEY_DATE] = kpiBudget.Date;
                    dtRow[TableDailyBudget.KEY_AVDELING] = main.appConfig.Avdeling;
                    dtRow[TableDailyBudget.KEY_BUDGET_ID] = budgetId;

                    dtRow[TableDailyBudget.KEY_BUDGET_TYPE] = element.type;
                    dtRow[TableDailyBudget.KEY_BUDGET_SALES] = element.Sales;
                    dtRow[TableDailyBudget.KEY_BUDGET_GM] = element.GM;
                    dtRow[TableDailyBudget.KEY_BUDGET_GM_PERCENT] = element.Percent;

                    table.Rows.Add(dtRow);
                }

                if (table != null && table.Rows.Count > 0)
                {
                    main.database.tableDailyBudget.RemoveDate(main.appConfig.Avdeling, kpiBudget.Date);

                    Log.d("Lagrer budsjett (id " + budgetId + ") med dato "
                        + kpiBudget.Date.ToShortDateString() + " til databasen..");
                    main.database.DoBulkCopy(table, TableDailyBudget.TABLE_NAME);
                    return true;
                }
                Log.d("Budsjett ble ikke lagret til databasen. Mangler viktig data");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return false;
        }

        private bool HeaderMatch(MessageHeader header)
        {
            return header != null && header.Subject != null
                && header.Subject.ToLower().Contains(matchSubjectMustContain)
                && header.From.HasValidMailAddress;
        }

        private bool AttachmentMatch(MessagePart attachment)
        {
            return attachment != null
                && attachment.FileName.ToLower().Contains(matchAttachmentFilenameMustContain)
                && attachment.FileName.ToLower().EndsWith(".pdf");
        }

        private KpiBudget ParseC810(byte[] bytestream)
        {
            try
            {
                PdfReader reader = new PdfReader(bytestream);

                string content = System.Text.Encoding.UTF8.GetString(reader.GetPageContent(1));

                int indexTotal = FindIndex(content, 0, "[(Total)]TJ");
                if (!VerifyIndex(indexTotal, "Fant ikke først element 'Total'"))
                    return null;

                int indexCategory = FindIndex(content, indexTotal + 1, "[(Category)]TJ");
                if (!VerifyIndex(indexCategory, "Fant ikke andre element 'Category'"))
                    return null;

                int indexSale = FindIndex(content, indexCategory + 1, "[(Sale/ Budget incl. VAT - Month to Date )]TJ");
                if (!VerifyIndex(indexSale, "Fant ikke tredje element 'Sale/ Budget incl. VAT - Month to Date'"))
                    return null;

                //int indexTotal = content.IndexOf("[(Total)]TJ");
                //int indexCategory = content.IndexOf("[(Category)]TJ", indexTotal + 1);
                //int indexSale = content.IndexOf("[(Sale/ Budget incl. VAT - Month to Date )]TJ", indexCategory + 1);

                string strBudget = content.Substring(indexCategory, indexSale - indexCategory);

                List<string> lines = strBudget.Split('\n').ToList();

                lines.RemoveAll(item => !item.ToString().StartsWith("[("));

                for (int i = lines.Count - 1; i >= 0; i--)
                    lines[i] = lines[i].Trim().Replace("[(", string.Empty).Replace(")]TJ", string.Empty).Replace(" ", string.Empty).Replace("/", string.Empty);

                var budget = new KpiBudget();
                budget.Date = selectedDate;

                for (int i = 0; i < lines.Count; i++)
                {
                    if (lines[i].Contains("MDA") || lines[i].Contains("AudioVideo") || lines[i].Contains("SDA") || lines[i].Contains("Telecom")
                        || lines[i].Contains("Computing") || lines[i].Contains("Kitchen") || lines[i].Contains("Other") || lines[i].Contains("Total"))
                    {
                        string type = lines[i];
                        i++;
                        decimal decSales = 0;
                        decimal.TryParse(lines[i], out decSales);
                        i++;
                        decimal decGM = 0;
                        decimal.TryParse(lines[i], out decGM);
                        var element = new KpiBudgetElement();
                        element.Insert(type, decSales, decGM);
                        budget.element.Add(element);
                    }
                }

                return budget;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return null;
        }

        private int FindIndex(string content, int start, string first, string second = "", string third = "")
        {
            try
            {
                int index = content.IndexOf(first, start);
                if (index > -1)
                    return index;
                if (second != null && second.Length > 0)
                    index = content.IndexOf(second, start);
                if (index > -1)
                    return index;
                if (third != null && third.Length > 0)
                    index = content.IndexOf(third, start);
            }
            catch (Exception ex)
            {
                Log.d("FindIndex unntak: " + ex.Message);
            }
            return -1;
        }

        private bool VerifyIndex(int index, string errorMsg, int minimum = 0, string minErrorMsg = "")
        {
            if (index == -1)
            {
                Log.d(errorMsg);
                return false;
            }
            else if (index < minimum)
            {
                Log.d(minErrorMsg);
                return false;
            }
            return true;
        }
    }

    public class KpiBudget
    {
        public DateTime Date { get; set; }
        public List<KpiBudgetElement> element { get; set; }

        public KpiBudget()
        {
            element = new List<KpiBudgetElement> { };
        }
    }

    public class KpiBudgetElement
    {
        public string type { get; set; }
        public decimal Sales { get; set; }
        public decimal GM { get; set; }
        public decimal Percent { get; set; }

        public KpiBudgetElement() { }
        public void Insert(string typeArg, decimal salesArg, decimal gmArg)
        {
            type = typeArg;
            Sales = salesArg;
            GM = gmArg;

            try
            {
                if (Sales != 0)
                    Percent = Math.Round(((GM * 1.25M) / Sales) * 100, 1);
                else
                    Percent = 0;
            }
            catch
            {
                Percent = 0;
            }
        }
    }
}
