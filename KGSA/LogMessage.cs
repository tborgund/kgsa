using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public class Logg
    {
        private static List<KgsaLog> log = new List<KgsaLog>();
        private static Color colorBtServer = Color.FromArgb(22, 93, 120);
        private static Color colorBtServerError = Color.FromArgb(120, 84, 22);

        public static event EventHandler LogAdded;

        /// <summary>
        /// Send logg melding
        /// </summary>
        /// <param name="message">Selve meldingen i string format</param>
        /// <param name="color">Fargen på meldingen</param>
        /// <param name="sendToLogOnly">Send bare til loggen</param>
        /// <param name="statusOnly">Send bare til status feltet</param>
        /// <param name="sendToFileOnly">Send bare til fil</param>
        public static void Log(string message, Color? color = null, bool sendToLogOnly = false, bool statusOnly = false, bool sendToFileOnly = false, bool markedAsDebug = false)
        {
            if (markedAsDebug)
                Console.WriteLine("Debug: " + message);

            var logObject = new KgsaLog(message, color, sendToLogOnly, statusOnly, sendToFileOnly, markedAsDebug);

            log.Add(logObject);

            if (LogAdded != null)
                LogAdded(null, EventArgs.Empty);
        }

        public static void BtServer(string message, bool error = false)
        {
            if (error)
                Log("Bluetooth server: " + message, colorBtServerError, true);
            else
                Log("Bluetooth server: " + message, colorBtServer, true);
        }

        public static void WebUser(string msg, HttpListenerRequest req, Color? c = null)
        {
            if (c == null)
                c = Color.DarkGoldenrod;
            Log("Web (" + req.UserHostAddress + "): " + msg, c);
        }

        public static void DebugSql(string msg)
        {
            Log(msg, Color.BlueViolet, true, false, false, true);
        }

        public static void Debug(string msg)
        {
            Log(msg, Color.Brown, true, false, false, true);
        }

        public static DialogResult Alert(string txt)
        {
            return Alert(txt, "KGSA", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Alert(string txt, string title)
        {
            return Alert(txt, title, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Alert(string txt, string title, MessageBoxButtons msgButton)
        {
            return Alert(txt, title, msgButton, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Alert(string txt, string title, MessageBoxButtons msgButton, MessageBoxIcon msgIcon)
        {
            return Alert(txt, title, msgButton, msgIcon, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Alert(string txt, string title, MessageBoxButtons msgButton, MessageBoxIcon msgIcon, MessageBoxDefaultButton msgDefaultButton)
        {
            if (msgIcon == MessageBoxIcon.Exclamation)
                Log(txt, Color.Red);
            else
                Log(txt, null, true);
            return MessageBox.Show(txt, title, msgButton, msgIcon, msgDefaultButton);
        }

        public static void Debug(string msg, Exception ex)
        {
            Log(msg + Environment.NewLine + "Exception melding: " + ex.Message + Environment.NewLine + "Exception: " + ex.ToString(), Color.Brown, true, false, false, true);
        }

        public static void Status(string msg, Color? c = null)
        {
            Log(msg, c, false, true);
        }

        public static void Unhandled(Exception ex, bool dialog = false)
        {
            Log("Uhåndtert unntak oppstod! Unntak melding: " + ex.Message + Environment.NewLine + "Exception: " + ex.ToString(), Color.Red);
            if (dialog)
                Alert("Ooops!\nNoe uventet skjedde som egentlig ikke skulle skje.\nMelding: " + ex.Message + "\nSe logg for detaljer.\n\nPrøv igjen, hvis det ikke fungerer; start programmet på nytt.", "Noe galt har skjedd", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        /// <summary>
        /// Show detailed modal error dialog
        /// </summary>
        /// <param name="exception">Exception object, optionally Null, but not recommended</param>
        /// <param name="details">Explain the possible nature of the problem when possible</param>
        /// <param name="title">The title of the error dialg</param>
        public static void ErrorDialog(Exception exception, string details, string title)
        {
            Log("Uhåndtert unntak oppstod! Unntak melding: " + exception.Message + Environment.NewLine + "Exception: " + exception.ToString(), Color.Red);
            using (FormError error = new FormError(title, exception, details))
            {
                error.ShowDialog();
            }
        }

        public static KgsaLog GetLastLog()
        {
            if (log.Count > 0)
                return log[log.Count - 1];
            else
                return null;
        }
    }

    public class KgsaLog
    {
        public string message { get; set; }
        public Color color { get; set; }
        public bool logonly { get; set; }
        public bool statusonly { get; set; }
        public bool fileonly { get; set; }
        public bool debug { get; set; }
        public KgsaLog(string msg, Color? c = null, bool log = false, bool status = false, bool file = false, bool debugArg = false)
        {
            this.message = msg;
            this.color = c.HasValue ? c.Value : Color.Black;
            this.logonly = log;
            this.statusonly = status;
            this.fileonly = file;

            if (this.statusonly && this.logonly)
                this.statusonly = false;
            this.debug = debugArg;
        }

    }
}
