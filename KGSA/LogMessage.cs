using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public class Log
    {
        private static List<KgsaLog> log = new List<KgsaLog>();
        private static Color colorBtServer = Color.FromArgb(22, 93, 120);
        private static Color colorBtServerError = Color.FromArgb(120, 84, 22);

        public static event EventHandler LogAdded;

        /// <summary>
        /// Create a log entry
        /// </summary>
        /// <param name="message">Log message</param>
        /// <param name="color">Color of the message</param>
        /// <param name="onlyShowInLog">Show in Log window only</param>
        /// <param name="onlyShowInStatus">Send bare til status feltet</param>
        /// <param name="onlyShowInLogFile">Send bare til fil</param>
        public static void n(string message, Color? color = null, bool onlyShowInLog = false, bool onlyShowInStatus = false, bool onlyShowInLogFile = false, bool isDebug = false, bool isSqlDebug = false)
        {
            if (isDebug || isSqlDebug)
                Console.WriteLine("Debug: " + message);

            var logObject = new KgsaLog(message, color, onlyShowInLog, onlyShowInStatus, onlyShowInLogFile, isDebug, isSqlDebug);

            log.Add(logObject);

            if (LogAdded != null)
                LogAdded(null, EventArgs.Empty);
        }

        public static void e(string message, bool showOnlyInLog = false)
        {
            n(message, Color.Red, showOnlyInLog);
        }

        public static void v(string message)
        {
            n(message, null, true, false, true);
        }

        public static void d(string message, Exception exception)
        {
            n(message + Environment.NewLine + "Exception: " + exception.Message + Environment.NewLine + "StackTrace: " + exception.StackTrace, Color.Brown, true, false, false, true);
        }

        public static void d(string message)
        {
            n(message, Color.Brown, true, false, false, true);
        }

        public static void BtServer(string message, bool error = false)
        {
            if (error)
                n("Bluetooth server: " + message, colorBtServerError, true);
            else
                n("Bluetooth server: " + message, colorBtServer, true);
        }

        public static void WebUser(string msg, HttpListenerRequest req, Color? c = null)
        {
            if (c == null)
                c = Color.DarkGoldenrod;
            n("Web (" + req.UserHostAddress + "): " + msg, c);
        }

        public static void DebugSql(string msg)
        {
            n(msg, Color.BlueViolet, true, false, false, true, true);
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
                n("Dialog: " + txt, Color.Red);
            else
                n("Dialog: " + txt, null, true);
            return MessageBox.Show(txt, title, msgButton, msgIcon, msgDefaultButton);
        }

        public static void Status(string msg, Color? c = null)
        {
            n(msg, c, false, true);
        }

        public static void Unhandled(Exception ex, bool dialog = false)
        {
            n("Uhåndtert unntak oppstod! Unntak melding: " + ex.Message + Environment.NewLine + "Exception: " + ex.ToString(), Color.Brown, false, false, false, true);
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
            n("Uhåndtert unntak oppstod! Unntak melding: " + exception.Message + Environment.NewLine + "Exception: " + exception.ToString(), Color.Red);
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
        public bool debugSql { get; set; }
        public KgsaLog(string msg, Color? c = null, bool log = false, bool status = false, bool file = false, bool debugArg = false, bool debugSqlArg = false)
        {
            this.message = msg;
            this.color = c.HasValue ? c.Value : Color.Black;
            this.logonly = log;
            this.statusonly = status;
            this.fileonly = file;
            if (this.statusonly && this.logonly)
                this.statusonly = false;
            this.debug = debugArg;
            this.debugSql = debugSqlArg;
        }

    }
}
