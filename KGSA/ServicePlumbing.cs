using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KGSA
{
    public partial class Service
    {
        public void MakeTableHeaderRapport(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Service nr</td>");
            doc.Add("<th class=\"{sorter: 'date'}\" width=50 >Mottat</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=50 >Status</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=50 >Dager</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Navn</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=80 >Selgerkode</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Verksted</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=50 >Behandlet</td>");
            doc.Add("</tr></thead>");
        }

        private string ColorDager(int dager)
        {
            if (dager > 21)
                return "<span style='color:red'>" + dager + "</span>";
            if (dager > 14)
                return "<span style='color:orange'>" + dager + "</span>";
            if (dager >= 0)
                return "<span style='color:green'>" + dager + "</span>";

            return dager.ToString();
        }

        public DataTable ReadyTableGraphAdvanced()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Index", typeof(int));
            dataTable.Columns.Add("Mottat", typeof(DateTime));
            dataTable.Columns.Add("Store", typeof(StorageService));
            return dataTable;
        }

        private string TatTall(string arg)
        {
            try
            {
                if (arg == "")
                    return main.appConfig.visningNull;

                decimal var = Math.Round(Convert.ToDecimal(arg), 2);
                string value = main.appConfig.visningNull;
                if (var < 0)
                    value = "<span style='color:red'>" + var.ToString("0.00") + "</span>";
                if (var > 0)
                    value = var.ToString("0.00");
                return value;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return main.appConfig.visningNull;
            }
        }

        public string Percent(string arg)
        {
            try
            {
                double var = Convert.ToDouble(arg);
                if (var == -1)
                    return "0&nbsp;%";
                if (var > 999)
                    var = 999;
                if (var < -999)
                    var = -999;
                string value = Math.Round(var, 2).ToString();
                return value + "&nbsp;%";
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return main.appConfig.visningNull;
            }
        }

        public bool ClearDatabase()
        {
            try
            {
                DialogResult msgbox = MessageBox.Show("Sikker på at du vil slette alle servicer fra databasen?",
                    "KGSA - VIKTIG",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);

                if (msgbox == DialogResult.Yes)
                {
                    Logg.Log("Tømmer service databasen..");

                    main.database.tableService.Reset();
                    main.database.tableServiceLogg.Reset();

                    dbServiceDatoFra = FormMain.rangeMin;
                    dbServiceDatoTil = FormMain.rangeMin;

                    Logg.Log("Service databasen nullstilt.", Color.Green);
                    return true;
                }
            }
            catch (Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod ved sletting av service databasen.", ex);
                errorMsg.ShowDialog();
            }
            return false;
        }

        public DataTable ReadyTableService()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("OrdreNr", typeof(int));
            dataTable.Columns.Add("DatoMottat", typeof(DateTime));
            dataTable.Columns.Add("Status", typeof(string));
            dataTable.Columns.Add("Dager", typeof(int));
            dataTable.Columns.Add("Navn", typeof(string));
            dataTable.Columns.Add("Selgerkode", typeof(string));
            dataTable.Columns.Add("Verksted", typeof(string));
            dataTable.Columns.Add("Egenservice", typeof(bool));
            dataTable.Columns.Add("ServiceID", typeof(int));
            dataTable.Columns.Add("FerdigBehandlet", typeof(bool));
            return dataTable;
        }

        private float CalcPercent(float fra, float mot)
        {
            try
            {
                float percent = 0;
                if (mot != 0)
                    percent = fra / mot;
                else
                    percent = 0;
                if (float.IsInfinity(percent))
                    percent = 1;
                if (float.IsNaN(percent))
                    percent = 0;
                if (percent < 0)
                    percent = 0;
                return percent;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return 0;
            }
        }

        private void RunPreLoadData()
        {
            try
            {
                string dateFormat = "dd/MM/yyyy HH:mm:ss";
                this.dbServiceDatoFra = FormMain.rangeMin;
                this.dbServiceDatoTil = FormMain.rangeMin;
                this.dbServiceDato = FormMain.rangeMin;

                SqlCeCommand cmd = new SqlCeCommand("SELECT MIN(DatoMottat) AS Expr1 FROM tblService WHERE (Avdeling = '" + main.appConfig.Avdeling + "')", main.connection);
                string temp = cmd.ExecuteScalar().ToString();
                if (temp != "")
                    this.dbServiceDatoFra = DateTime.ParseExact(temp, dateFormat, FormMain.norway);

                cmd = new SqlCeCommand("SELECT MAX(DatoMottat) AS Expr1 FROM tblService WHERE (Avdeling = '" + main.appConfig.Avdeling + "')", main.connection);
                temp = cmd.ExecuteScalar().ToString();
                if (temp != "")
                    this.dbServiceDatoTil = DateTime.ParseExact(temp, dateFormat, FormMain.norway);

                cmd = new SqlCeCommand("SELECT MAX(DatoTid) AS Expr1 FROM tblServiceLogg", main.connection);
                temp = cmd.ExecuteScalar().ToString();
                if (temp != "")
                    this.dbServiceDato = DateTime.ParseExact(temp, dateFormat, FormMain.norway);
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private Color Tint(Color source, Color tint, decimal alpha)
        {
            //(tint -source)*alpha + source
            int red = Convert.ToInt32(((tint.R - source.R) * alpha + source.R));
            int blue = Convert.ToInt32(((tint.B - source.B) * alpha + source.B));
            int green = Convert.ToInt32(((tint.G - source.G) * alpha + source.G));
            return Color.FromArgb(255, red, green, blue);
        }
    }

    public class ServiceHistory
    {
        public int totalt { get; set; }
        public int aktive { get; set; }
        public decimal tat { get; set; }
        public decimal over14prosent { get; set; }
        public decimal over21prosent { get; set; }
        public int tilarbeid { get; set; }
        public ServiceHistory(int totaltArg, int aktiveArg, decimal tatArg, decimal over14prosentArg, decimal over21prosentArg, int tilarbeidArg)
        {
            this.totalt = totaltArg;
            this.aktive = aktiveArg;
            this.tat = tatArg;
            this.over14prosent = over14prosentArg;
            this.over21prosent = over21prosentArg;
            this.tilarbeid = tilarbeidArg;
        }
    }


    public class StorageService
    {
        public List<StorageServieArray> servicer { get; set; }
        public int antall { get; set; }
        public StorageService()
        {
            servicer = new List<StorageServieArray> { };
        }
    }

    public class StorageServieArray
    {
        public int status { get; set; }
        public string selgerkode { get; set; }
        public DateTime Iarbeid { get; set; }
        public DateTime Ferdig { get; set; }
        public DateTime Utlevert { get; set; }
        public int tat { get; set; }
        public int tilarbeid { get; set; }
        public int ordrenr { get; set; }
        public bool egenservice { get; set; }
        public bool ferdigbehandlet { get; set; }

        public StorageServieArray(int statusArg, int tatArg, int ordrenrArg, string selgerkodeArg, bool egenserviceArg, DateTime IarbeidArg, int tilarbeidArg, DateTime FerdigArg, DateTime UtlevertArg, bool Ferdigbehandlet)
        {
            this.egenservice = egenserviceArg;
            this.selgerkode = selgerkodeArg;
            this.status = statusArg;
            this.Iarbeid = IarbeidArg;
            this.Ferdig = FerdigArg;
            this.Utlevert = UtlevertArg;
            this.ordrenr = ordrenrArg;
            this.tat = tatArg;
            this.tilarbeid = tilarbeidArg;
            this.ferdigbehandlet = Ferdigbehandlet;
        }

        public StorageServieArray(int statusArg, int tatArg, int ordrenrArg, string selgerkodeArg, bool egenserviceArg, DateTime IarbeidArg, int tilarbeidArg, DateTime FerdigArg, DateTime UtlevertArg)
        {
            this.egenservice = egenserviceArg;
            this.selgerkode = selgerkodeArg;
            this.status = statusArg;
            this.Iarbeid = IarbeidArg;
            this.Ferdig = FerdigArg;
            this.Utlevert = UtlevertArg;
            this.ordrenr = ordrenrArg;
            this.tat = tatArg;
            this.tilarbeid = tilarbeidArg;
        }

        public StorageServieArray(int statusArg, int tatArg, int ordrenrArg, string selgerkodeArg, bool egenserviceArg, DateTime IarbeidArg, int tilarbeidArg, DateTime FerdigArg)
        {
            this.egenservice = egenserviceArg;
            this.selgerkode = selgerkodeArg;
            this.status = statusArg;
            this.Iarbeid = IarbeidArg;
            this.Ferdig = FerdigArg;
            this.Utlevert = FormMain.rangeMin;
            this.ordrenr = ordrenrArg;
            this.tat = tatArg;
            this.tilarbeid = tilarbeidArg;
        }

        public StorageServieArray(int statusArg, int tatArg, int ordrenrArg, string selgerkodeArg, bool egenserviceArg, DateTime IarbeidArg, int tilarbeidArg)
        {
            this.egenservice = egenserviceArg;
            this.selgerkode = selgerkodeArg;
            this.status = statusArg;
            this.Iarbeid = IarbeidArg;
            this.Ferdig = FormMain.rangeMin;
            this.Utlevert = FormMain.rangeMin;
            this.ordrenr = ordrenrArg;
            this.tat = tatArg;
            this.tilarbeid = tilarbeidArg;
        }
        public StorageServieArray(int statusArg, int tatArg, int ordrenrArg, string selgerkodeArg, bool egenserviceArg)
        {
            this.egenservice = egenserviceArg;
            this.selgerkode = selgerkodeArg;
            this.status = statusArg;
            this.Iarbeid = FormMain.rangeMin;
            this.Ferdig = FormMain.rangeMin;
            this.Utlevert = FormMain.rangeMin;
            this.ordrenr = ordrenrArg;
            this.tat = tatArg;
        }
    }

    public class ServiceStatus
    {
        public string VenterService = "Venter service";
        public string IarbeidInternt = "I arb. internt";
        public string IarbeidExternt = "I arb. eksternt";
        public string VenterLev = "venter svar fra lev.";
        public string VenterRfc = "Venter svar RFC";
        public string VenterFor = "venter svar forsikr.";
        public string VenterInn = "Venter Innsending";
        public string VenterKun = "venter Kunderespons";

        public string VenterInnsending = "Venter Innsending"; // NY!
        public string VenterDeler = "Venter deler"; // NY!
        public string MottatPcKlinikk = "Mottat PC-Klinikk"; // NY!
        public string IarbeidPcKlinikk = "I Arbeid PC-Klinikk"; // NY!
        public string DeleventPcKlinikk = "Delevent PC-Klinikk"; // NY!

        public string Ferdig = "Ferdig/venter utlev";
        public string FerdigKlinikk = "Ferdig PC-Klinikk";
        public string FerdigUtlevert = "Ferdig, utlevert";

        public int GetStatusInt(string statusArg)
        {
            if (statusArg == VenterService)
                return 1;
            if (statusArg == IarbeidInternt)
                return 2;
            if (statusArg == VenterLev)
                return 3;
            if (statusArg == VenterRfc)
                return 4;
            if (statusArg == VenterFor)
                return 5;
            if (statusArg == VenterInn)
                return 6;
            if (statusArg == VenterKun)
                return 7;

            if (statusArg == VenterInnsending)
                return 8;
            if (statusArg == VenterDeler)
                return 8;
            if (statusArg == VenterDeler)
                return 8;
            if (statusArg == MottatPcKlinikk)
                return 9;
            if (statusArg == IarbeidPcKlinikk)
                return 9;
            if (statusArg == DeleventPcKlinikk)
                return 9;


            if (statusArg == IarbeidExternt)
                return 10;
            if (statusArg == Ferdig)
                return 90;
            if (statusArg == FerdigKlinikk)
                return 91;
            if (statusArg == FerdigUtlevert)
                return 99;
            return 0;
        }

        public Color GetStatusColor(int statusArg)
        {
            if (statusArg == 1)
                return Color.Red;
            if (statusArg == 2)
                return Color.Indigo;
            if (statusArg == 3)
                return Color.Orchid;
            if (statusArg == 4)
                return Color.Orchid;
            if (statusArg == 5)
                return Color.Orchid;
            if (statusArg == 6)
                return Color.Red;
            if (statusArg == 7)
                return Color.Indigo;
            if (statusArg == 8)
                return Color.Olive;
            if (statusArg == 9)
                return Color.Navy;
            if (statusArg == 10)
                return Color.Brown;
            if (statusArg == 90)
                return System.Drawing.ColorTranslator.FromHtml("#3daa3d");
            if (statusArg == 91)
                return System.Drawing.ColorTranslator.FromHtml("#3daa8c");
            if (statusArg == 99)
                return Color.LightGreen;

            return Color.LightGray;
        }
    }

    /// <summary>
    /// Calculates a moving average value over a specified window.  The window size must be specified
    /// upon creation of this object.
    /// </summary>
    /// <remarks>Authored by Drew Noakes, February 2005.  Use freely, though keep this message intact and
    /// report any bugs to me.  I also appreciate seeing extensions, or simply hearing that you're using
    /// these classes.  You may not copyright this work, though may use it in commercial/copyrighted works.
    /// Happy coding.
    ///
    /// Updated 29 March 2007.  Added a Reset() method.</remarks>
    public sealed class MovingAverageCalculator
    {
        private readonly int _windowSize;
        private readonly float[] _values;
        private int _nextValueIndex;
        private float _sum;
        private int _valuesIn;

        /// <summary>
        /// Create a new moving average calculator.
        /// </summary>
        /// <param name="windowSize">The maximum number of values to be considered
        /// by this moving average calculation.</param>
        /// <exception cref="ArgumentOutOfRangeException">If windowSize less than one.</exception>
        public MovingAverageCalculator(int windowSize)
        {
            if (windowSize < 1)
                throw new ArgumentOutOfRangeException("windowSize", windowSize, "Window size must be greater than zero.");

            _windowSize = windowSize;
            _values = new float[_windowSize];

            Reset();
        }

        /// <summary>
        /// Updates the moving average with its next value, and returns the updated average value.
        /// When IsMature is true and NextValue is called, a previous value will 'fall out' of the
        /// moving average.
        /// </summary>
        /// <param name="nextValue">The next value to be considered within the moving average.</param>
        /// <returns>The updated moving average value.</returns>
        /// <exception cref="ArgumentOutOfRangeException">If nextValue is equal to float.NaN.</exception>
        public float NextValue(float nextValue)
        {
            if (float.IsNaN(nextValue))
                throw new ArgumentOutOfRangeException("nextValue", "NaN may not be provided as the next value.  It would corrupt the state of the calculation.");

            // add new value to the sum
            _sum += nextValue;

            if (_valuesIn < _windowSize)
            {
                // we haven't yet filled our window
                _valuesIn++;
            }
            else
            {
                // remove oldest value from sum
                _sum -= _values[_nextValueIndex];
            }

            // store the value
            _values[_nextValueIndex] = nextValue;

            // progress the next value index pointer
            _nextValueIndex++;
            if (_nextValueIndex == _windowSize)
                _nextValueIndex = 0;

            return _sum / _valuesIn;
        }

        /// <summary>
        /// Gets a value indicating whether enough values have been provided to fill the
        /// speicified window size.  Values returned from NextValue may still be used prior
        /// to IsMature returning true, however such values are not subject to the intended
        /// smoothing effect of the moving average's window size.
        /// </summary>
        public bool IsMature
        {
            get { return _valuesIn == _windowSize; }
        }

        /// <summary>
        /// Clears any accumulated state and resets the calculator to its initial configuration.
        /// Calling this method is the equivalent of creating a new instance.
        /// </summary>
        public void Reset()
        {
            _nextValueIndex = 0;
            _sum = 0;
            _valuesIn = 0;
        }
    }
}
