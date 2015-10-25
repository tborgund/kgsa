using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using KGSA.Properties;
using System.Threading;
using System.Text;
using System.Collections;
using System.Globalization;

namespace KGSA
{
    public class Ranking
    {
        public CultureInfo norway = new CultureInfo("nb-NO");
        public bool velgerPeriode = true;
        public List<string> vkButikk = new List<string>();
        public List<string> vkButikkAlle = new List<string>();
        public Avdeling avdeling = new Avdeling();
        public DataTable dt;
        public DataTable dtDay;
        public DataTable dtMonth;
        public DataTable dtAvd;
        public DataTable dtQuick;
        public DataTable dtCompare;
        public DataTable dtCompareLastMonth;
        public DataTable dtLastWeek;
        public DataTable sqlce = new DataTable();
        public DateTime dtFra = DateTime.Now;
        public DateTime dtTil = DateTime.Now;
        public DateTime dtPick = DateTime.Now;
        public string strFra = "";
        public string strTil = "";
        public string[] favoritter;
        public Random random = new Random();
        public string outerclass = "OutertableNormal";
        public AutoResetEvent tblReadyMonth = new AutoResetEvent(false);
        public AutoResetEvent tblReadyDay = new AutoResetEvent(false);
        public AutoResetEvent tblReadyFav = new AutoResetEvent(false);
        public AutoResetEvent tblReadyCompare = new AutoResetEvent(false);
        public AutoResetEvent tblReadyCompareLastMonth = new AutoResetEvent(false);
        public AutoResetEvent tblReadyQuick = new AutoResetEvent(false);
        public bool provisjon = false;
        public FormMain main;

        public Ranking()
        {
            favoritter = FormMain.Favoritter.ToArray();
        }

        internal string ObjectToClassStr(object obj)
        {
            if (DBNull.Value == obj)
                return "";

            bool complete = (bool)obj;
            if (complete)
                return "-complete";
            else
                return "";

        }

        public int ObjectToInteger(object obj)
        {
            try
            {
                if (DBNull.Value == obj)
                    return 0;
                return Convert.ToInt32(obj);
            }
            catch (Exception) { }
            return 0;
        }

        public string ObjectToDateTimeShortString(object obj)
        {
            try
            {
                if (DBNull.Value == obj)
                    return "";
                return Convert.ToDateTime(obj).ToShortDateString();
            }
            catch (Exception) { }
            return "";
        }

        public static bool IsOdd(int value)
        {
            return value % 2 != 0;
        }

        public static DateTime GetFirstDayOfMonth(DateTime givenDate)
        {
            return new DateTime(givenDate.Year, givenDate.Month, 1);
        }

        public static DateTime GetLastDayOfMonth(DateTime givenDate)
        {
            return GetFirstDayOfMonth(givenDate).AddMonths(1).Subtract(new TimeSpan(1, 0, 0, 0, 0));
        }

        internal decimal Compute(DataTable table, string search, string filter = null)
        {
            object r = table.Compute(search, filter);
            if (!DBNull.Value.Equals(r))
                return Convert.ToDecimal(r);
            else
                return 0;
        }

        internal string GetSqlStringFor(string katArg)
        {
            if (katArg == "MDA")
                return "(Varegruppe >= 100 AND Varegruppe < 200) ";
            if (katArg == "AudioVideo")
                return "(Varegruppe >= 200 AND Varegruppe < 300) ";
            if (katArg == "SDA")
                return "(Varegruppe >= 300 AND Varegruppe < 400) ";
            if (katArg == "Tele")
                return "(Varegruppe >= 400 AND Varegruppe < 500) ";
            if (katArg == "Data")
                return "(Varegruppe >= 500 AND Varegruppe < 600) ";
            else
                return " ";
        }

        public int GetIntAvdelingFor(string katArg)
        {
            if (katArg == "MDA")
                return 1;
            if (katArg == "AudioVideo")
                return 2;
            if (katArg == "SDA")
                return 3;
            if (katArg == "Tele")
                return 4;
            if (katArg == "Data")
                return 5;
            else
                return -1;
        }

        public DataTable ValidateDataTable(DataTable dt)
        {
            try
            {
                if (dt.Columns[2].DataType == typeof(string))
                {
                    DataTable dtCloned = dt.Clone();
                    dtCloned.Columns[2].DataType = typeof(Int16);
                    foreach (DataRow row in dt.Rows)
                        dtCloned.ImportRow(row);

                    return dtCloned;
                }
                else
                    return dt;
            }
            catch { return dt; };
        }

        public DateTime FindLastTransactionday(DateTime from, DateTime to)
        {
            try
            {
                DataTable dt = main.database.GetSqlDataTable("SELECT TOP(1) Dato FROM tblSalg WHERE Avdeling = " + main.appConfig.Avdeling
                    + " AND CONVERT(NVARCHAR(10),Dato,121) >= CONVERT(NVARCHAR(10),'" + from.ToString("yyyy-MM-dd")
                    + "',121) AND CONVERT(NVARCHAR(10),Dato,121) <= CONVERT(NVARCHAR(10),'" + to.ToString("yyyy-MM-dd") + "',121)"
                    + " ORDER BY Dato DESC");
                if (dt.Rows.Count > 0)
                {
                    DateTime output = Convert.ToDateTime(dt.Rows[0]["Dato"]);
                    return output;
                }
            }
            catch(Exception ex)
            {
                Log.Unhandled(ex);
            }
            return DateTime.Now;
        }

        public DataTable RemoveDuplicateRows(DataTable dTable, string colName)
        {
            Hashtable hTable = new Hashtable();
            ArrayList duplicateList = new ArrayList();

            //Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            //And add duplicate item value in arraylist.
            foreach (DataRow drow in dTable.Rows)
            {
                if (hTable.Contains(drow[colName]))
                    duplicateList.Add(drow);
                else
                    hTable.Add(drow[colName], string.Empty);
            }

            //Removing a list of duplicate items from datatable.
            foreach (DataRow dRow in duplicateList)
                dTable.Rows.Remove(dRow);

            //Datatable which contains unique records will be return as output.
            return dTable;
        }

        public bool ErKat(string str)
        {
            if (str == "MDA" || str == "Data" || str == "SDA" || str == "Tele" || str == "AudioVideo" || str == "Kasse" || str == "Aftersales" || str == "TOTALT" || str == "Telecom" || str == "Computing" || str == "Cross" || str == "Kjøkken")
                return true;
            else
                return false;
        }
        public bool ErKatSpecial(string str)
        {
            if (str == "Kasse" || str == "Aftersales" || str == "Cross")
                return true;
            else
                return false;
        }

        public DataTable ReadyTableButikk()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Kat", typeof(string));
            dataTable.Columns.Add("Salg", typeof(int));
            dataTable.Columns.Add("Omset", typeof(decimal));
            dataTable.Columns.Add("Inntjen", typeof(decimal));
            dataTable.Columns.Add("OmsetExMva", typeof(decimal));
            //for (int i = 0; i < vkButikk.Count; i++)
            //    dataTable.Columns.Add("VK_" + vkButikk[i], typeof(int));
            dataTable.Columns.Add("Prosent", typeof(double));
            dataTable.Columns.Add("AntallTjen", typeof(int));
            dataTable.Columns.Add("TjenOmset", typeof(decimal));
            dataTable.Columns.Add("TjenInntjen", typeof(decimal));
            dataTable.Columns.Add("TjenMargin", typeof(double));
            dataTable.Columns.Add("StromInntjen", typeof(decimal));
            dataTable.Columns.Add("StromAntall", typeof(int));
            dataTable.Columns.Add("StromMargin", typeof(double));
            dataTable.Columns.Add("ModInntjen", typeof(decimal));
            dataTable.Columns.Add("ModOmset", typeof(decimal));
            dataTable.Columns.Add("ModAntall", typeof(int));
            dataTable.Columns.Add("ModMargin", typeof(double));
            dataTable.Columns.Add("FinansInntjen", typeof(decimal));
            dataTable.Columns.Add("FinansAntall", typeof(int));
            dataTable.Columns.Add("FinansMargin", typeof(double));

            dataTable.Columns.Add("Kuppvarer", typeof(int));
            dataTable.Columns.Add("AccessoriesAntall", typeof(int));
            dataTable.Columns.Add("AccessoriesInntjen", typeof(double));
            dataTable.Columns.Add("AccessoriesOmset", typeof(double));
            dataTable.Columns.Add("AccessoriesMargin", typeof(double)); // f36565
            dataTable.Columns.Add("AccessoriesSoB", typeof(double)); // f36565
            dataTable.Columns.Add("SnittAntall", typeof(int));
            dataTable.Columns.Add("SnittInntjen", typeof(double));
            dataTable.Columns.Add("SnittOmset", typeof(double));
            dataTable.Columns.Add("SnittOmsetAlle", typeof(double));
            dataTable.Columns.Add("TjenHitrate", typeof(double));
            
            return dataTable;
        }

        public void MakeTableHeaderBudsjett(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=149 >Kategori</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=95 >Omsetn.</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=95 >Inntjen.</td>");

            doc.Add("<th colspan=2 class=\"{sorter: 'digit'}\" width=90 style='background:#f5954e;'>Finans (SoM)</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#f5954e;'><abbr title='Btokr. inntjen. Finans / Btokr inntjen. alle varer'>%</abbr></td>");

            doc.Add("<th colspan=2 class=\"{sorter: 'digit'}\" width=90 style='background:#6699ff;'>TA (SoM)</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#6699ff;'><abbr title='Inntjen. TA / Btokr. inntjen. alle varer'>%</abbr></td>");

            doc.Add("<th colspan=2 class=\"{sorter: 'digit'}\" width=90 style='background:#FAF39E;'>Strøm (SoM)</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#FAF39E;'><abbr title='Btokr. inntjen. Strøm / Btokr. inntjen. alle varer'>%</abbr></td>");

            doc.Add("<th colspan=2 class=\"{sorter: 'digit'}\" width=90 style='background:#80c34a;'>RTG/SA (SoM)</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'><abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>%</abbr></td>");

            doc.Add("</tr></thead>");
        }

        public void MakeTableHeaderTrivia(List<string> doc)
        {
            doc.Add("<thead><tr>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=120 >Kategori</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=95 >Inntjen.</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=95 >Omset.</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=40 >#</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=85 >Bilagsnr.</td>");

            doc.Add("</tr></thead>");
        }

        public bool StopRankingPending()
        {
            if (FormMain.stopRanking)
                return true;
            return false;
        }

        public string Forkort(string source)
        {
            int tail_length =  main.appConfig.kolForkortLengde;
            if ( main.appConfig.kolForkort)
            {
                if (tail_length >= source.Length)
                    return source;
                source = "<abbr class='tips' title='" + source + "'>.." + source.Substring(source.Length - tail_length) + "</abbr>";
                return source;
            }
            else
            {
                return source;
            }
        }

        public string ForkortTekst(string source, int max)
        {
            try
            {
                if (source.Length > (max + 2))
                    return source.Substring(0, max) + "..";
            }
            catch (Exception) { };
            return source;
        }

        public decimal CalcHitrate(decimal serv, decimal comp)
        {
            decimal hit = 0;
            try
            {
                if (comp < 0 && serv < 0)
                    return -1;
                if (comp > 0 && serv > 0)
                    return Math.Round(serv / comp * 100, 2);
                if (comp <= 0 && serv > 0)
                    return 100;
                if (comp <= 0 && serv < 0)
                    return 0;
                if (comp <= 0 && serv != 0)
                    return 100;
                if (serv == 0 && comp == 0)
                    return -1;
                else if (serv == 0 && comp != 0)
                    return 0;
            }
            catch
            {
                hit = 0;
            }
            return hit;
        }

        public string ForkortTall(decimal arg)
        {
            try
            {
                if (arg == 0)
                    return  main.appConfig.visningNull;

                var r = Math.Round((int)arg / 1000d, 1);
                if (r < 98)
                    return arg.ToString("#,##0");

                var b = Math.Round((int)arg / 1000000d, 1);
                if (b < 2)
                    return r.ToString("#,##0") + " k";

                return b.ToString() + " m";
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                return "";
            }
        }

        public string PlusMinus(object arg)
        {
            return PlusMinus(Convert.ToString(arg));
        }

        public string PlusMinus(object arg, bool green)
        {
            return PlusMinus(Convert.ToString(arg), green);
        }

        public string PlusMinus(int arg)
        {
            return PlusMinus(Convert.ToString(arg));
        }

        public string PlusMinus(decimal arg)
        {
            return PlusMinus(Convert.ToString(arg));
        }

        public string PlusMinus(double arg)
        {
            return PlusMinus(Convert.ToString(arg));
        }

        public string PlusMinus(string arg, bool green = false)
        {
            try
            {
                if (String.IsNullOrEmpty(arg))
                    return  main.appConfig.visningNull;

                decimal var = Math.Round(Convert.ToDecimal(arg), 0);
                string value =  main.appConfig.visningNull;
                if (var < 0)
                    value = "<span style='color:red'>" + var.ToString("#,##0") + "</span>";
                if (var > 0)
                {
                    if (green)
                        value = "<span style='color:green'>" + var.ToString("#,##0") + "</span>";
                    else
                        value = var.ToString("#,##0");
                }

                return value;
            }
            catch
            {
                return  main.appConfig.visningNull;
            }
        }


        public string Number(decimal arg)
        {
            try
            {
                if (arg == 0)
                    return  main.appConfig.visningNull;

                decimal var = Math.Round(arg, 0);
                string value =  main.appConfig.visningNull;
                if (var < 0)
                    value = "<span style='color:red;'>" + var.ToString("#,##0") + "</span>";
                if (var > 0)
                    value = "<span style='color:green;'>+" + var.ToString("#,##0") + "</span>";
                return value;
            }
            catch
            {
                return  main.appConfig.visningNull;
            }
        }

        public string NumberPercent(decimal arg, int deci = 2)
        {
            try
            {
                if (arg == 0)
                    return  main.appConfig.visningNull;

                decimal var = Math.Round(arg, deci);
                string value =  main.appConfig.visningNull;
                if (var < 0)
                    value = "<span style='color:red'>" + var + " %</span>";
                if (var > 0)
                    value = "<span style='color:green'>+" + var + " %</span>";
                return value;
            }
            catch
            {
                return  main.appConfig.visningNull;
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
                string value = Math.Round(var, 0).ToString();
                return value + "&nbsp;%";
            }
            catch
            {
                return  main.appConfig.visningNull;
            }
        }

        public string PercentDecimalToColoredString(decimal arg)
        {
            try
            {
                decimal var = Convert.ToDecimal(arg);
                string value = "";
                if (var > 99 || var < -99)
                    value = Math.Round(var, 0).ToString();
                else if (var < -9)
                    value = Math.Round(var, 1).ToString();
                else
                    value = Math.Round(var, 2).ToString();
                if (var == 0)
                    return  main.appConfig.visningNull;
                if (var > 999)
                    value = "999";
                if (var < -999)
                    value = "-999";

                if (var < 0)
                    return "<span style='color:red'>" + value + "&nbsp;%</span>";
                if (var > 0)
                    return value + "&nbsp;%";
                return  main.appConfig.visningNull;
            }
            catch
            {
                return  main.appConfig.visningNull;
            }
        }

        public string PercentShare(object arg)
        {
            return PercentShare(Convert.ToString(arg));
        }

        public string PercentShare(string arg, string argTjen = "", bool kat = false)
        {
            try
            {
                if (String.IsNullOrEmpty(arg))
                    arg = "0";
                decimal var = Convert.ToDecimal(arg);
                string value = "";
                if (var > 99 || var < -99)
                    value = Math.Round(var, 0).ToString();
                else if (var < -9)
                    value = Math.Round(var, 1).ToString();
                else
                    value = Math.Round(var, 2).ToString();
                if (var == 0)
                    return  main.appConfig.visningNull;
                if (var > 999)
                    value = "999";
                if (var < -999)
                    value = "-999";

                if (var < 0)
                    return "<span style='color:red'>" + value + "&nbsp;%</span>";
                if (var > 0)
                    return value + "&nbsp;%";
                return  main.appConfig.visningNull;
            }
            catch
            {
                return  main.appConfig.visningNull;
            }
        }

        public string PercentStyleData(string arg)
        {
            try
            {
                decimal var = Convert.ToDecimal(arg);
                if (var == -1)
                    return "";
                string output = "color:";
                output += ColorTranslator.ToHtml(SelectColorData(var, true));
                output += ";background:";
                output += ColorTranslator.ToHtml(SelectColorData(var, false));
                output += ";";
                return output;
            }
            catch
            {
                return "";
            }
        }

        public string PercentStyleNett(string arg)
        {
            try
            {
                decimal var = Convert.ToDecimal(arg);
                if (var == -1)
                    return "";
                string output = "color:";
                output += ColorTranslator.ToHtml(SelectColorNett(var, true));
                output += ";background:";
                output += ColorTranslator.ToHtml(SelectColorNett(var, false));
                output += ";";
                return output;
            }
            catch
            {
                return "";
            }
        }

        public Color SelectColorData(decimal var, bool font)
        {
            Color color = Color.White;
            if (var <=  main.appConfig.color1max)
            {
                var value = (var /  main.appConfig.color1max);
                color = Tint( main.appConfig.color1,  main.appConfig.color2, value);
            }
            else if (var >=  main.appConfig.color2min && var <=  main.appConfig.color2max)
            {
                var value = (var -  main.appConfig.color2min) / ( main.appConfig.color2max -  main.appConfig.color2min);
                color = Tint( main.appConfig.color2,  main.appConfig.color3, value);
            }
            else if (var >=  main.appConfig.color3min && var <=  main.appConfig.color3max)
            {
                var value = (var -  main.appConfig.color3min) / ( main.appConfig.color3max -  main.appConfig.color3min);
                color = Tint( main.appConfig.color3,  main.appConfig.color4, value);
            }
            else if (var >=  main.appConfig.color4min && var <=  main.appConfig.color4max)
            {
                var value = (var -  main.appConfig.color4min) / ( main.appConfig.color4max -  main.appConfig.color4min);
                color = Tint( main.appConfig.color4,  main.appConfig.color5, value);
            }
            else if (var >=  main.appConfig.color5min && var <=  main.appConfig.color5max)
            {
                color =  main.appConfig.color5;
            }
            else if (var >=  main.appConfig.color6min)
            {
                color =  main.appConfig.color6;
            }
            if (!font)
                return color;

            var inverter = true;
            Color c = color;
            var l = 0.333 * c.R + 0.333 * c.G + 0.333 * c.B;
            if (l > 100)
                inverter = false;

            Color fontcolor = Color.Black;
            if (var <=  main.appConfig.color1max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.color2min && var <=  main.appConfig.color2max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.color3min && var <=  main.appConfig.color3max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.color4min && var <=  main.appConfig.color4max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.color5min && var <=  main.appConfig.color5max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.color6min && inverter)
            {
                fontcolor = Color.White;
            }
            return fontcolor;
        }

        private Color Tint(Color source, Color tint, decimal alpha)
        {
            try
            {
                if ( main.appConfig.visningJevnfarge)
                {
                    //(tint -source)*alpha + source
                    int red = Convert.ToInt32(((tint.R - source.R) * alpha + source.R));
                    int blue = Convert.ToInt32(((tint.B - source.B) * alpha + source.B));
                    int green = Convert.ToInt32(((tint.G - source.G) * alpha + source.G));
                    return Color.FromArgb(255, red, green, blue);
                }
                else
                {
                    return source;
                }
            }
            catch
            {
                return Color.Red;
            }
        }

        public Color SelectColorNett(decimal var, bool font)
        {
            Color color = Color.White;
            if (var <=  main.appConfig.ncolor1max)
            {
                var value = (var /  main.appConfig.ncolor1max);
                color = Tint( main.appConfig.ncolor1,  main.appConfig.ncolor2, value);
            }
            else if (var >=  main.appConfig.ncolor2min && var <=  main.appConfig.ncolor2max)
            {
                var value = (var -  main.appConfig.ncolor2min) / ( main.appConfig.ncolor2max -  main.appConfig.ncolor2min);
                color = Tint( main.appConfig.ncolor2,  main.appConfig.ncolor3, value);
            }
            else if (var >=  main.appConfig.ncolor3min && var <=  main.appConfig.ncolor3max)
            {
                var value = (var -  main.appConfig.ncolor3min) / ( main.appConfig.ncolor3max -  main.appConfig.ncolor3min);
                color = Tint( main.appConfig.ncolor3,  main.appConfig.ncolor4, value);
            }
            else if (var >=  main.appConfig.ncolor4min && var <=  main.appConfig.ncolor4max)
            {
                var value = (var -  main.appConfig.ncolor4min) / ( main.appConfig.ncolor4max -  main.appConfig.ncolor4min);
                color = Tint( main.appConfig.ncolor4,  main.appConfig.ncolor5, value);
            }
            else if (var >=  main.appConfig.ncolor5min && var <=  main.appConfig.ncolor5max)
            {
                color =  main.appConfig.ncolor5;
            }
            else if (var >=  main.appConfig.ncolor6min)
            {
                color =  main.appConfig.ncolor6;
            }
            if (!font)
                return color;

            var inverter = true;
            Color c = color;
            var l = 0.333 * c.R + 0.333 * c.G + 0.333 * c.B;
            if (l > 100)
                inverter = false;

            Color fontcolor = Color.Black;
            if (var <=  main.appConfig.ncolor1max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.ncolor2min && var <=  main.appConfig.ncolor2max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.ncolor3min && var <=  main.appConfig.ncolor3max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.ncolor4min && var <=  main.appConfig.ncolor4max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.ncolor5min && var <=  main.appConfig.ncolor5max && inverter)
            {
                fontcolor = Color.White;
            }
            else if (var >=  main.appConfig.ncolor6min && inverter)
            {
                fontcolor = Color.White;
            }
            return fontcolor;
        }
    }
}