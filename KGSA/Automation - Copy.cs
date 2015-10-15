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
    public partial class AutoQuick : Form
    {
        public AppSettings appConfig = new AppSettings();
        public string message = "";
        private Timer tm;
        private List<string> program;
        public DateTime Dato { get; set; }
        public int errorCode { get; set; }
        public int macroAttempt { get; set; }
        // errorCode forklaring:
        // 0 = alt OK
        // 1 = udefinert
        // 2 = avbrutt av bruker
        // 5 = feil under lasting av makro program

        public AutoQuick()
        {
            InitializeComponent();
            LoadSettings();
            macroAttempt = 0;
            try
            {
                string input = File.ReadAllText(Form1.macroProgram);
                program = new List<string>(input.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries));
            }
            catch
            {
                message = "Feil under lasting av macro.";
                errorCode = 5;
            }

            tm = new Timer();
            tm.Interval = 2 * 1000;
            tm.Tick += new EventHandler(tm_Tick);
        }

        private void Automation_Load(object sender, EventArgs e)
        {
            startMacro();
        }

        public void LoadSettings()
        {
            if (!File.Exists(Form1.settingsFile))
            {
                return;
            }

            XmlSerializer mySerializer = new XmlSerializer(typeof(AppSettings));
            using (StreamReader myXmlReader = new StreamReader(Form1.settingsFile))
            {
                try
                {
                    appConfig = (AppSettings)mySerializer.Deserialize(myXmlReader);
                }
                catch { }
            }
        }

        private string GetActiveWindowTitle()
        {
            const int nChars = 256;
            IntPtr handle = IntPtr.Zero;
            StringBuilder Buff = new StringBuilder(nChars);
            handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString().ToLower();
            }
            return "";
        }

        void ActivateApp(string processName)
        {
            Process[] p = Process.GetProcessesByName(processName);
            if (p.Count() > 0)
                SetForegroundWindow(p[0].MainWindowHandle);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            startMacro();
        }

        private void startMacro()
        {
            if (!bgWorkerMacro.IsBusy)
            {
                button6.Text = "Stop";
                progressBar.Style = ProgressBarStyle.Marquee;
                bgWorkerMacro.RunWorkerAsync();
            }
            else
            {
                Message("Avbryter..");
                bgWorkerMacro.CancelAsync();
                Message("Bruker stoppet makro.", Color.Red);
                errorCode = 2;
            }
        }

        delegate void SetTextCallback(string str, Color? c = null);

        public void Message(string str, Color? c = null)
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
                    this.Invoke(d, new object[] { str, c });
                }
                else
                {
                    labelLog.Text = str;
                    labelLog.ForeColor = c.HasValue ? c.Value : Color.Black;
                    message = str;
                }
            }
            catch
            {
                labelLog.Text = "Feil i beskjeds system.";
                labelLog.ForeColor = Color.Red;
                message = str;
            }
        }

        private void Unlock()
        {
            if (appConfig.unlock && appConfig.unlockArg != "")
            {
                try
                {
                    var Logon = new ProcessStartInfo();
                    Logon.FileName = Form1.pluginLogin;
                    Logon.Arguments = appConfig.unlockArg;
                    Logon.WorkingDirectory = Form1.settingsPath;
                    Logon.UseShellExecute = false;
                    Logon.RedirectStandardOutput = true;
                    Logon.CreateNoWindow = true;
                    Process L = Process.Start(Logon);
                    while (!L.StandardOutput.EndOfStream)
                    {
                        if (L.StandardOutput.ReadLine().Length > 2)
                            Message("Unlock: " + L.StandardOutput.ReadLine());
                    }
                    L.WaitForExit();
                    System.Threading.Thread.Sleep(2000);
                    L.Dispose();
                }
                catch
                {
                    // Ikke gjør noe.
                }
            }
        }

        private void bgWorkerMacro_DoWork(object sender, DoWorkEventArgs e)
        {
            if (program.Count > 0)
            {
                Unlock();

                if (macroAttempt == 3)
                {
                    try
                    {
                        System.Diagnostics.Process[] IEProcesses = System.Diagnostics.Process.GetProcessesByName("telnet98.exe");
                        foreach (System.Diagnostics.Process CurrentProcess in IEProcesses)
                        {
                            if (CurrentProcess.MainWindowTitle.Contains("elguide"))
                            {
                                CurrentProcess.Kill();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Message("Unntak oppstod ved avslutting av prosesser - Exception message: " + ex.ToString());
                    }
                }

                for (int i = 0; i < program.Count; i++)
                {
                    if (program[i].ToLower().Contains("{dato}"))
                    {
                        if (Dato.Date.Equals(DateTime.Now.Date))
                            Dato = Form1.GetFirstDayOfMonth(Dato);
                        program[i] = program[i].Replace("{dato}", Dato.ToString("ddMMyy"));
                        program[i] = program[i].Replace("{Dato}", Dato.ToString("ddMMyy"));
                    }
                }

                int latency = Convert.ToInt32((1000 * appConfig.macroLatency) + (macroAttempt * 300)); // beregn forsinkelse
                double span = (DateTime.Now - Dato).TotalDays;
                int extraWait = 0;
                if (span > 31)
                    Dato = DateTime.Now.AddMonths(-1); // Begrens oss til å importere en måned bak i tid
                if (span > 3)
                    extraWait = Convert.ToInt32(span * 10); // Legg til 10 ekstra sekunder pr. dag.    

                
                int wait = 0;

                try
                {
                    System.Diagnostics.Process.Start(appConfig.macroElguide);

                    try
                    {
                        for (int i = 0; i < program.Count; i++)
                        {
                            if (bgWorkerMacro.CancellationPending)
                                break;

                            string linje = program[i];

                            if (linje.StartsWith("FindProcess"))
                            {
                                string process = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value;
                                ActivateApp(process);
                                Message("[" + i + "] Finner prosess: " + process);

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            if (linje.StartsWith("WaitForTitle"))
                            {
                                string title = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value.ToLower();

                                Message("[" + i + "] Venter på vindu med tekst: " + title);

                                int teller = 0;
                                while (!GetActiveWindowTitle().Contains(title) && teller < 30)
                                {
                                    System.Threading.Thread.Sleep(latency);
                                    teller++;
                                    if (bgWorkerMacro.CancellationPending)
                                        break;
                                }
                                if (teller >= 30)
                                {
                                    Message("Feil: Ventet for lenge på vindu med tittel " + title + ".");
                                    break;
                                }

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            else if (linje.StartsWith("Wait"))
                            {
                                var value = Convert.ToInt32(Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value);
                                if (value > 10)
                                    wait = extraWait;
                                else
                                    wait = 0;
                                double delay = value + wait;
                                
                                int teller = Convert.ToInt32(Math.Round(delay * appConfig.macroLatency, 0));

                                while(teller > 0)
                                {
                                    Message("[" + i + "] Vent i " + teller + " sekunder.");
                                    System.Threading.Thread.Sleep(1000);
                                    teller--;
                                    if (bgWorkerMacro.CancellationPending)
                                        break;
                                }
                            }
                            else if (linje.StartsWith("KeyHoldStart"))
                            {
                                string key = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value;

                                Message("[" + i + "] KeyHoldStart: " + key);

                                if (key == "SHIFT")
                                    InputSimulator.SimulateKeyDown(VirtualKeyCode.SHIFT);
                                if (key == "CONTROL")
                                    InputSimulator.SimulateKeyDown(VirtualKeyCode.CONTROL);

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            else if (linje.StartsWith("KeyHoldEnd"))
                            {
                                string key = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value;

                                Message("[" + i + "] KeyHoldEnd: " + key);

                                if (key == "SHIFT")
                                    InputSimulator.SimulateKeyUp(VirtualKeyCode.SHIFT);
                                if (key == "CONTROL")
                                    InputSimulator.SimulateKeyUp(VirtualKeyCode.CONTROL);

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            else if (linje.StartsWith("KeyPress"))
                            {
                                string key = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value;

                                Keys key_restored = (Keys)Enum.Parse(typeof(Keys), key);

                                Message("[" + i + "] KeyPress: " + key);
                                InputSimulator.SimulateKeyPress((VirtualKeyCode)key_restored);

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            else if (linje.StartsWith("KeyString"))
                            {
                                string str = Regex.Match(linje, @"\(([^)]*)\)").Groups[1].Value;

                                Message("[" + i + "] KeyString: " + str);

                                InputSimulator.SimulateTextEntry(str);

                                System.Threading.Thread.Sleep(latency / 2);
                            }
                            if (i + 1 == program.Count)
                            {
                                Message("Fullført.");
                                errorCode = 0;
                            }
                        }
                    }
                    catch
                    {
                        Message("Ukjent feil oppstod.");
                    }
                }
                catch
                {
                    Message("Feil: Fant ikke elguide profil eller ingen tilgang.");
                }
            }
        }

        private void bgWorkerMacro_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string str = e.UserState.ToString();
            Message(str);
        }

        private void bgWorkerMacro_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button6.Text = "Kjør";
            progressBar.Style = ProgressBarStyle.Continuous;
            tm.Start();
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

        private void Automation_FormClosing(object sender, FormClosingEventArgs e)
        {
            bgWorkerMacro.CancelAsync();
            Message("Bruker stoppet makro.", Color.Red);
        }
    }
}
