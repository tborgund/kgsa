using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Diagnostics;
using WindowsInput;
using System.Text.RegularExpressions;

namespace KGSA
{
    public partial class FormMacro : Form
    {
        public AppSettings appConfig;
        private Timer tm;
        private List<string> program;
        private bool ignoreExtraWait;
        delegate void SetTextCallback(string str, Color? c = null, bool nosave = false);
        BackgroundWorker bwMaster;
        BackgroundWorker bwMacro = new BackgroundWorker();

        public DataTable tableQuick = getTable();
        public KveldstallInfo KveldstallInfo = new KveldstallInfo();
        public DateTime Dato { get; set; }
        public int errorCode { get; set; }
        // errorCode forklaring:
        // 0 = alt OK
        // 1 = udefinert feil
        // 2 = avbrutt av bruker
        // 3 = feil som er definert men som ikke skal avbryte videre program kjøring
        // 5 = feil under lasting av makro program
        // 6 = kritisk feil, må avbrytes
        public string errorMessage { get; set; }
        public int macroAttempt { get; set; }

        public FormMacro(AppSettings app, DateTime dateArg, string programArg, int macroAttemptArg, bool argIgnoreExtraWait, BackgroundWorker bwArg)
        {
            this.bwMaster = bwArg;
            this.appConfig = app;
            this.Dato = dateArg;
            this.macroAttempt = macroAttemptArg;
            this.ignoreExtraWait = argIgnoreExtraWait;
            errorCode = 0;
            errorMessage = "";

            try
            {
                string input = File.ReadAllText(programArg);
                program = new List<string>(input.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries));
                FilterProgram();
            }
            catch (Exception ex)
            {
                Error(6, "Feil under lasting av makro program.", ex);
            }

            InitializeComponent();
            if (appConfig.macroShowWarning)
                panelWarning.Visible = true;
            else
                this.Height = 95;

            tm = new Timer();
            tm.Interval = 2 * 1000;
            tm.Tick += new EventHandler(tm_Tick);

            bwMacro.DoWork += new DoWorkEventHandler(bwMacro_DoWork);
            bwMacro.ProgressChanged += new ProgressChangedEventHandler(bwMacro_ProgressChanged);
            bwMacro.WorkerReportsProgress = true;
            bwMacro.WorkerSupportsCancellation = true;
            bwMacro.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwMacro_RunWorkerCompleted);
            bwMacro.RunWorkerAsync();
        }

        private void bwMacro_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (macroAttempt >= 3)
                    KillElguideProcesses();

                int latency = Convert.ToInt32((1000 * appConfig.macroLatency) + (macroAttempt * 300)); // beregn forsinkelse
                int extraWait = 0;
                double span = (DateTime.Now - Dato).TotalDays;
                if (span > 3 && !ignoreExtraWait)
                    extraWait = Convert.ToInt32(span * 5); // Legg til 5 ekstra sekunder pr. dag.    
                if (extraWait > 150)
                    extraWait = 150; // Maks 150 ekstra sekunder
                int wait = 0;

                try
                {
                    for (int i = 0; i < program.Count; i++)
                    {
                        if (bwMacro.CancellationPending)
                            break;

                        if (bwMaster != null)
                            if (bwMaster.CancellationPending)
                                break;

                        string programLine = program[i];

                        if (programLine.StartsWith("Start"))
                        {
                            string str = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Message("[" + i + "] Starter applikasjon: " + str + "..");

                            if (File.Exists(str))
                                System.Diagnostics.Process.Start(str);
                            else
                            {
                                Error(6, "Kunne ikke starte (" + str + ") - Fil mangler!");
                                return;
                            }
                            System.Threading.Thread.Sleep(latency);
                        }
                        else if (programLine.StartsWith("FindRow"))
                        {
                            string[] a = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value.Split(';');
                            Message("[" + i + "] Finner rad: " + a[1] + " file: " + a[0]);

                            if (!FindRow(a))
                            {
                                Error(1, "Fant ikke rad!");
                                break;
                            }
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("ImportNow"))
                        {
                            string strAvdeling = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Message("[" + i + "] Importerer data for " + strAvdeling + "..");

                            System.Threading.Thread.Sleep(2000);

                            if (0 != ImportCSV(Convert.ToInt32(strAvdeling)))
                            {
                                Error(1, "Feil oppstod under importering av CSV.");
                                return;
                            }
                        }
                        else if (programLine.StartsWith("ImportAllSales"))
                        {
                            string strAvdeling = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;

                            Message("[" + i + "] Importerer salgs data for " + strAvdeling + "..");

                            System.Threading.Thread.Sleep(2000);

                            if (0 != ImportAllSalesCSV(Convert.ToInt32(strAvdeling)))
                                Error(3, "Feil oppstod under importering av Salgs CSV.");
                        }
                        else if (programLine.StartsWith("FindProcess"))
                        {
                            string process = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            ActivateApp(process);
                            Message("[" + i + "] Finner prosess: " + process);

                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("WaitForTitle"))
                        {
                            string title = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value.ToLower();

                            Message("[" + i + "] Venter på vindu med tekst: " + title);

                            int teller = 0;
                            while (!GetActiveWindowTitle().Contains(title) && teller < 30)
                            {
                                System.Threading.Thread.Sleep(latency);
                                teller++;
                                if (bwMacro.CancellationPending)
                                    break;
                                if (bwMaster != null)
                                    if (bwMaster.CancellationPending)
                                        break;
                            }
                            if (teller >= 30)
                            {
                                Error(1, "Ventet for lenge på vindu med tittel '" + title + "'");
                                break;
                            }
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("Wait"))
                        {
                            var value = Convert.ToInt32(Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value);
                            if (value > 10)
                                wait = extraWait;
                            else
                                wait = 0;
                            decimal delay = value + wait;

                            var teller = Math.Round(delay * appConfig.macroLatency, 0);
                            Message("[" + i + "] Vent i " + teller + " sekunder.");

                            while (teller > 0)
                            {
                                if (teller > 0)
                                    Message("[" + i + "] Vent i " + teller + " sekunder.", null, true);
                                System.Threading.Thread.Sleep(1000);
                                teller--;
                                if (bwMacro.CancellationPending)
                                    break;
                                if (bwMaster != null)
                                    if (bwMaster.CancellationPending)
                                        break;
                            }
                        }
                        else if (programLine.StartsWith("KeyHoldStart"))
                        {
                            string key = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Message("[" + i + "] KeyHoldStart: " + key);

                            if (key == "SHIFT")
                                InputSimulator.SimulateKeyDown(VirtualKeyCode.SHIFT);
                            if (key == "CONTROL")
                                InputSimulator.SimulateKeyDown(VirtualKeyCode.CONTROL);
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("KeyHoldEnd"))
                        {
                            string key = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Message("[" + i + "] KeyHoldEnd: " + key);

                            if (key == "SHIFT")
                                InputSimulator.SimulateKeyUp(VirtualKeyCode.SHIFT);
                            if (key == "CONTROL")
                                InputSimulator.SimulateKeyUp(VirtualKeyCode.CONTROL);
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("KeyPress"))
                        {
                            string key = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Keys key_restored = (Keys)Enum.Parse(typeof(Keys), key);

                            Message("[" + i + "] KeyPress: " + key);

                            InputSimulator.SimulateKeyPress((VirtualKeyCode)key_restored);
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("KeyString"))
                        {
                            string str = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;

                            Message("Venter på Elguide vindu..");
                            int teller = 0;
                            while (!GetActiveWindowTitle().Contains("elguide") && teller < 30)
                            {
                                System.Threading.Thread.Sleep(latency);
                                teller++;
                                if (bwMacro.CancellationPending)
                                    break;
                                if (bwMaster != null)
                                    if (bwMaster.CancellationPending)
                                        break;
                            }
                            if (teller >= 30)
                            {
                                Error(1, "Ventet for lenge Elguide vindu.");
                                break;
                            }

                            Message("[" + i + "] KeyString: " + str + "..");

                            InputSimulator.SimulateTextEntry(str);
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("Password"))
                        {
                            string str = Regex.Match(programLine, @"\(([^)]*)\)").Groups[1].Value;
                            Message("[" + i + "] Password");

                            string password = str;

                            if (Regex.IsMatch(str, @"^\d+$") && str.Length > 10)
                            {
                                SimpleAES simple = new SimpleAES();
                                password = simple.DecryptString(str);
                            }
                            
                            if (!GetActiveWindowTitle().Contains("elguide"))
                            {
                                Error(6, "Forventet vindu var IKKE aktivt!");
                                System.Threading.Thread.Sleep(latency / 2);
                                break;
                            }

                            InputSimulator.SimulateTextEntry(password);
                            System.Threading.Thread.Sleep(latency / 2);
                        }
                        else if (programLine.StartsWith("ImportKveldstall"))
                        {
                            Message("[" + i + "] Legger til kveldstall kommandoer..");

                            GenerateCommandsQuick();
                            System.Threading.Thread.Sleep(latency / 2);
                        }

                        if (i + 1 == program.Count)
                        {
                            Message("Makro fullført!", Color.Green);
                            errorCode = 0;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Error(6, "Unntak oppstod under makro kjøring.", ex);
                    FormError errorMsg = new FormError("Unntak oppstod under makro kjøring", ex);
                    errorMsg.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Error(1, "Ukjent feil", ex);
            }
        }

        private void bwMacro_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string str = e.UserState.ToString();
            Message(str);
        }

        private void bwMacro_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
                if (bwMaster != null)
                    bwMaster.CancelAsync();
            progressBar.Style = ProgressBarStyle.Continuous;
            tm.Start();
        }

        private void GenerateCommandsQuick()
        {
            for (int f = 0; f < FormMain.Favoritter.Count; f++)
            {
                program.Add("Wait(3)");
                program.Add("KeyString(136)");
                program.Add("KeyPress(Enter)");
                program.Add("KeyString(" + FormMain.Favoritter[f] + ")");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(3)");
                program.Add("KeyPress(Enter)");
                program.Add("KeyPress(Enter)");
                program.Add("KeyPress(F12)");
                program.Add("KeyPress(Right)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(10)");
                program.Add("KeyPress(F1)");
                program.Add("KeyPress(F1)");
                program.Add("KeyPress(F1)");
                program.Add("Wait(4)");
                program.Add("ImportNow(" + FormMain.Favoritter[f] + ")");
            }

            if (appConfig.macroImportQuickSales)
            {
                program.Add("Wait(3)");
                program.Add("KeyString(136)");
                program.Add("KeyPress(Enter)");
                program.Add("KeyString(" + FormMain.Favoritter[0] + ")");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(3)");
                program.Add("KeyPress(Enter)");
                program.Add("KeyPress(Enter)"); // Vi er i oversikts menyen

                program.Add("KeyPress(Down)"); // Marker Lyd og bilde
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyPress(F12)");
                program.Add("KeyPress(Right)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyString(inegoAV)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(10)");
                program.Add("KeyPress(F1)");
                program.Add("Wait(2)");

                program.Add("KeyPress(Down)"); // Marker SDA
                program.Add("KeyPress(Down)"); // Marker Telecom
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyPress(F12)");
                program.Add("KeyPress(Right)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyString(inegoTele)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(10)");
                program.Add("KeyPress(F1)");
                program.Add("Wait(2)");

                program.Add("KeyPress(Down)"); // Marker Computing
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyPress(F12)");
                program.Add("KeyPress(Right)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(2)");
                program.Add("KeyString(inegoComp)");
                program.Add("KeyPress(Enter)");
                program.Add("Wait(10)");
                program.Add("KeyPress(F1)");
                program.Add("KeyPress(F1)");
                program.Add("KeyPress(F1)");
                program.Add("KeyPress(F1)");
                program.Add("Wait(4)");
                program.Add("ImportAllSales(" + FormMain.Favoritter[0] + ")");
            }

            // program for avslutting av elguide..
            program.Add("KeyPress(F1)");
            program.Add("KeyPress(F1)");
        }

        private bool FindRow(string[] array)
        {
            try
            {
                string file = array[0];
                string row = array[1];

                DataTable dt = ImportRows(appConfig.csvElguideExportFolder + @"\" + file.ToLower() + ".csv");

                int rows = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    rows++;
                    if (dt.Rows[i][6].ToString().Contains(row))
                        break;
                }

                if (rows > 0 && rows != dt.Rows.Count)
                {
                    rows++;
                    string key = "Down";
                    for (int i = 0; i < rows; i++)
                    {
                        Keys key_restored = (Keys)Enum.Parse(typeof(Keys), key);
                        Message("[X] KeyPress: " + key);
                        InputSimulator.SimulateKeyPress((VirtualKeyCode)key_restored);
                        int latency = Convert.ToInt32((1000 * appConfig.macroLatency) + (macroAttempt * 300)); // beregn forsinkelse
                        System.Threading.Thread.Sleep(latency / 2);

                        if (bwMaster != null)
                            if (bwMaster.CancellationPending)
                                break;
                    }
                    return true;
                }
                else
                {
                    Logg.Log("Auto: Fant ikke rad (" + row + ")");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                errorCode = 6;
                return false;
            }

        }

        private DataTable ImportRows(string filename)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    Message("Fant ikke CSV.", Color.Red);
                    errorCode = 6;
                    return null;
                }

                string[] Lines = File.ReadAllLines(filename);
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

                if (dt.Rows.Count > 0)
                    return dt;
                else
                {
                    Message("Fant ingen rader.");
                    Logg.Log("Auto: Fant ingen rader.", null, true);
                    return null;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private void FilterProgram()
        {
            bool missingStart = true;
            for (int i = 0; i < program.Count; i++)
            {
                if (program[i].Contains("{elguide}"))
                    program[i] = program[i].Replace("{elguide}", appConfig.macroElguide);
                if (program[i].Contains("{dato}"))
                    program[i] = program[i].Replace("{dato}", Dato.ToString("ddMMyy"));
                if (program[i].Contains("{Dato}"))
                    program[i] = program[i].Replace("{Dato}", Dato.ToString("ddMMyy"));
                if (program[i].Contains("{fradato}"))
                    program[i] = program[i].Replace("{fradato}", Dato.AddDays(-60).ToString("ddMMyy"));
                if (program[i].Contains("{tildato}"))
                    program[i] = program[i].Replace("{tildato}", Dato.ToString("ddMMyy"));
                if (program[i].Contains("{avdeling}"))
                    program[i] = program[i].Replace("{avdeling}", appConfig.Avdeling.ToString());
                if (program[i].Contains("{Avdeling}"))
                    program[i] = program[i].Replace("{Avdeling}", appConfig.Avdeling.ToString());
                if (program[i].StartsWith("Start"))
                    missingStart = false;
            }
            if (missingStart)
                program.Insert(0, "Start(" + appConfig.macroElguide + ")");
        }

        private int ImportCSV(int avdArg)
        {
            try
            {
                if (!File.Exists(appConfig.csvElguideExportFolder + "inego.csv"))
                {
                    Message("Fant ikke CSV.", Color.Red);
                    return 1;
                }

                else if (File.GetLastWriteTime(appConfig.csvElguideExportFolder + "inego.csv").Date != DateTime.Now.Date)
                {
                    Message("Eksport mislyktes, CSV ikke oppdatert!", Color.Red);
                    return 1;
                }

                string[] Lines = File.ReadAllLines(appConfig.csvElguideExportFolder + "inego.csv");
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

                if (dt.Rows.Count < 6)
                {
                    errorCode = 1;
                    errorMessage = "CSV fil var for kort. " + dt.Rows.Count + " linjer.";
                    return 1;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][0].ToString().Contains("---"))
                    {
                        // Vi er i row før totalen.
                        i++;
                        DataRow rowTot = tableQuick.NewRow();
                        rowTot["Favoritt"] = avdArg;
                        rowTot["Avdeling"] = "TOTAL";
                        rowTot["Salg"] = Convert.ToInt32(dt.Rows[i][1].ToString());
                        rowTot["Omsetn"] = Convert.ToDecimal(dt.Rows[i][2].ToString());
                        rowTot["Fritt"] = Convert.ToDecimal(dt.Rows[i][3].ToString());
                        rowTot["Fortjeneste"] = Convert.ToDecimal(dt.Rows[i][4].ToString());
                        rowTot["Margin"] = Convert.ToDouble(dt.Rows[i][5].ToString());
                        rowTot["Rabatt"] = Convert.ToDecimal(dt.Rows[i][6].ToString());
                        tableQuick.Rows.Add(rowTot);
                        break;
                    }

                    DataRow row = tableQuick.NewRow();
                    row["Favoritt"] = avdArg;
                    string tmp = dt.Rows[i][0].ToString().Substring(1, 1);
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
                    if (dt.Rows[i][1].ToString() != "")
                        row["Salg"] = Convert.ToInt32(dt.Rows[i][1].ToString());
                    else
                        row["Salg"] = 0;
                    if (dt.Rows[i][2].ToString() != "")
                        row["Omsetn"] = Convert.ToDecimal(dt.Rows[i][2].ToString());
                    else
                        row["Omsetn"] = 0;
                    if (dt.Rows[i][3].ToString() != "")
                        row["Fritt"] = Convert.ToDecimal(dt.Rows[i][3].ToString());
                    else
                        row["Fritt"] = 0;
                    if (dt.Rows[i][4].ToString() != "")
                        row["Fortjeneste"] = Convert.ToDecimal(dt.Rows[i][4].ToString());
                    else
                        row["Fortjeneste"] = 0;
                    if (dt.Rows[i][5].ToString() != "")
                        row["Margin"] = Convert.ToDouble(dt.Rows[i][5].ToString());
                    else
                        row["Margin"] = 0;
                    if (dt.Rows[i][6].ToString() != "")
                        row["Rabatt"] = Convert.ToDecimal(dt.Rows[i][6].ToString());
                    else
                        row["Rabatt"] = 0;
                    tableQuick.Rows.Add(row);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Message("Unntak i import funksjon. Exception: " + ex.ToString(), Color.Red);
                return 3;
            }
        }

        private int ImportAllSalesCSV(int avdArg)
        {
            try
            {
                string[] listOfFiles = new string[] { "inegoAV.csv", "inegoTele.csv", "inegoComp.csv" };

                foreach (string file in listOfFiles)
                {
                    if (!File.Exists(appConfig.csvElguideExportFolder + file))
                    {
                        Message("Fant ikke " + file, Color.Red);
                        return 1;
                    }
                    else if (File.GetLastWriteTime(appConfig.csvElguideExportFolder + file).Date != DateTime.Now.Date)
                    {
                        Message("Avbryter importering av " + file + " fordi filen er ikke oppdatert!", Color.Red);
                        return 1;
                    }

                    string[] Lines = File.ReadAllLines(appConfig.csvElguideExportFolder + file);
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
                        errorCode = 1;
                        errorMessage = "Sales CSV fil (" + file + ") var for kort. " + dt.Rows.Count + " linjer.";
                        return 1;
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        if (dt.Rows[i][0].ToString().Contains("580"))
                        {
                            KveldstallArray salg = new KveldstallArray();
                            salg.Type = "Computing";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            KveldstallInfo.Salg.Add(salg);
                            Logg.Debug("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }

                        if (dt.Rows[i][0].ToString().Contains("480"))
                        {
                            KveldstallArray salg = new KveldstallArray();
                            salg.Type = "Telecom";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            KveldstallInfo.Salg.Add(salg);
                            Logg.Debug("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }

                        if (dt.Rows[i][0].ToString().Contains("280"))
                        {
                            KveldstallArray salg = new KveldstallArray();
                            salg.Type = "AudioVideo";
                            salg.Antall = Convert.ToInt32(dt.Rows[i][1].ToString());
                            salg.Btokr = Convert.ToDouble(dt.Rows[i][4].ToString());
                            salg.Salgspris = Convert.ToDouble(dt.Rows[i][2].ToString());
                            KveldstallInfo.Salg.Add(salg);
                            Logg.Debug("Fant " + salg.Type + ": " + salg.Antall + " - " + salg.Btokr + " - " + salg.Salgspris);
                            continue;
                        }
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                Message("Unntak i import funksjon. Exception: " + ex.ToString(), Color.Red);
                return 3;
            }
        }

        private void tm_Tick(object sender, EventArgs e)
        {
            tm.Stop();
            this.Close();
        }

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lp1, string lp2);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private static DataTable getTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Favoritt", typeof(int));
            dt.Columns.Add("Avdeling", typeof(string));
            dt.Columns.Add("Salg", typeof(int));
            dt.Columns.Add("Omsetn", typeof(decimal));
            dt.Columns.Add("Fritt", typeof(decimal));
            dt.Columns.Add("Fortjeneste", typeof(decimal));
            dt.Columns.Add("Margin", typeof(double));
            dt.Columns.Add("Rabatt", typeof(decimal));
            return dt;
        }

        private string GetActiveWindowTitle()
        {
            try
            {
                const int nChars = 256;
                IntPtr handle = IntPtr.Zero;
                StringBuilder Buff = new StringBuilder(nChars);
                handle = GetForegroundWindow();

                if (GetWindowText(handle, Buff, nChars) > 0)
                    return Buff.ToString().ToLower();

                return "";
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }

        private void ActivateApp(string processName)
        {
            try
            {
                Process[] p = Process.GetProcessesByName(processName);
                if (p.Count() > 0)
                    SetForegroundWindow(p[0].MainWindowHandle);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void buttonStartStop_Click(object sender, EventArgs e)
        {
            stopMacro();
        }
        public void Message(string str, Color? c = null, bool nosave = false)
        {
            try
            {
                if (c == null)
                    c = Color.Black;
                str = str.Trim();
                str = str.Replace("\n", " ");
                str = str.Replace("\r", String.Empty);
                str = str.Replace("\t", String.Empty);
                if (str.ToLower().Contains("fullført"))
                    c = Color.Green;
                if (str.ToLower().Contains("feil"))
                    c = Color.Red;

                if (this.labelLog.InvokeRequired)
                {
                    SetTextCallback d = new SetTextCallback(Message);
                    this.Invoke(d, new object[] { str, c, nosave });
                }
                else
                {
                    if (!nosave)
                        Logg.Debug("Macro: " + str);
                    this.labelLog.Text = str;
                    this.labelLog.ForeColor = c.HasValue ? c.Value : Color.Black;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void Error(int code, string message, Exception ex = null)
        {
            errorCode = code;
            errorMessage = message;
            
            if (ex != null)
                Logg.Unhandled(ex);
            Logg.Log("Makro: " + message + " Kode: " + code, Color.Red);
            Message(message, Color.Red);
        }

        private void stopMacro()
        {
            bwMacro.CancelAsync();
            Message("Bruker stoppet makro.", Color.Red);
            int t = 0;
            while(bwMacro.IsBusy)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
                t++;
                if (t > 50)
                    Message("Venter på at alle operasjoner skal avsluttes..", Color.Red);
            }
            if (bwMaster != null)
                bwMaster.CancelAsync();
            
            errorCode = 2;
        }

        private void KillElguideProcesses()
        {
            try
            {

                System.Diagnostics.Process[] IEProcesses = System.Diagnostics.Process.GetProcessesByName("telnet98.exe");
                foreach (System.Diagnostics.Process CurrentProcess in IEProcesses)
                    if (CurrentProcess.MainWindowTitle.ToLower().Contains("elguide"))
                        CurrentProcess.Kill();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void FormMacro_FormClosing(object sender, FormClosingEventArgs e)
        {
            bwMacro.CancelAsync();
            if (bwMaster != null)
                if (bwMaster.CancellationPending)
                    errorCode = 2;
        }

    }

    public class KveldstallInfo
    {
        public int Avdeling;
        public DateTime Dato;
        public List<KveldstallArray> Salg { get; set; }
        public KveldstallInfo()
        {
            Salg = new List<KveldstallArray> { };
        }
    }

    public class KveldstallArray
    {
        public string Type;
        public int Antall;
        public double Btokr;
        public double Salgspris;
    }
}
