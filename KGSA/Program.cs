using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

namespace KGSA
{
    static class Program
    {
        /// <summary>
        /// Author: tborgund@gmail.com Trond Borgund
        /// Et program laget for å bistå presentasjon og ranking av salg av tjenester i Elkjøp.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (Mutex mutex = new Mutex(false, "Global\\" + appGuid))
            {
                if (!mutex.WaitOne(0, false))
                {
                    MessageBox.Show("KGSA programmet kjører allerede!");
                    return;
                }
                Application.Run(new FormMain(args));
            }
        }
        public static string appGuid = "5fbdf7ef-1969-48b3-a8a6-cdd98e40c2e2";
    }
}
