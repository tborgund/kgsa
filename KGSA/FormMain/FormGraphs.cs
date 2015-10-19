using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using KGSA.Properties;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.ComponentModel;
using System.Globalization;

namespace KGSA
{
    partial class FormMain
    {
        public static List<string> selgerkodeList = new List<string> { };
        string _graphKatCurrent = "Data";
        string _graphSelCurrent = "";
        int _graphLengde = -1;
        DateTime _graphFraDato = DateTime.Now;
        bool _graphInitialized = false;
        DateTime chkGraphPicker = rangeMin;
        public static bool _graphReqStop = false;

        private void clickTopGraph(MouseEventArgs e)
        {
            try
            {
                if (IsBusy() || EmptyDatabase() || topgraph == null)
                    return;

                DateTime d1 = topgraph.datoFra;
                DateTime d2 = topgraph.dato;

                int X = e.Location.X;
                float gWidth = graphPanelTop.Width;
                int days = (d2 - d1).Days;
                float Hstep = gWidth / days;
                int area = -1;
                for (int i = 0; i < gWidth; i++)
                {
                    if (X >= (i * Hstep) && X < (i * Hstep) + Hstep)
                        area = i;
                }
                int a = (days - area) - 1;
                if (area > -1 && area <= days && a > -1)
                {
                    DateTime d = d2.AddDays(-a);
                    if (d.Month == appConfig.dbTo.Month)
                    {
                        highlightDate = d;
                        pickerRankingDate.Value = d;
                        graphPanelTop.Invalidate();
                        tabControlMain.SelectedTab = tabPageRank;

                        var page = currentPage();
                        if (!String.IsNullOrEmpty(page))
                            RunRanking(page);
                        else
                            RunRanking("Data");
                    }
                }

            }
            catch(Exception ex)
            {
                Logg.Debug("Uhåndtert unntak oppstod ved clickTopGraph().", ex);
            }
        }

        private void UpdateGraphFields()
        {
            // Oppdater _graphSelCurrent som er gjeldene selger
            if (listBox_GraphSelgere.Items.Count > 0 && listBox_GraphSelgere.SelectedItem != null)
                _graphSelCurrent = listBox_GraphSelgere.SelectedItem.ToString();
            else
                _graphSelCurrent = "";
            if (_graphSelCurrent == "ALLE")
                _graphSelCurrent = "";

            // Oppdater _graphLengde som er gjeldene dager som skal vises
            if (comboBox_GraphLengde.SelectedIndex > -1)
                _graphLengde = comboBox_GraphLengde.SelectedIndex;

            if (datoPeriodeVelger)
            {
                _graphFraDato = datoPeriodeFra;
                pickerDato_Graph.Value = datoPeriodeTil;
            }
            else
            {
                // Oppdater fra dato
                if (_graphLengde > -1 && pickerDato_Graph.Value != null)
                {
                    if (_graphLengde == 0)
                        _graphFraDato = pickerDato_Graph.Value.AddMonths(-1);
                    if (_graphLengde == 1)
                        _graphFraDato = pickerDato_Graph.Value.AddMonths(-2);
                    if (_graphLengde == 2)
                        _graphFraDato = pickerDato_Graph.Value.AddMonths(-3);
                    if (_graphLengde == 3)
                        _graphFraDato = pickerDato_Graph.Value.AddMonths(-6);
                    if (_graphLengde == 4)
                        _graphFraDato = pickerDato_Graph.Value.AddYears(-1);
                    if (_graphLengde == 5)
                        _graphFraDato = pickerDato_Graph.Value.AddYears(-2);
                }
                else if (pickerDato_Graph.Value != null)
                    _graphFraDato = pickerDato_Graph.Value.AddMonths(-1);
                else
                    _graphFraDato = appConfig.dbTo.AddDays(-30);
            }
        }

        private void PaintGraph()
        {
            if (!EmptyDatabase() && gc != null && _graphInitialized)
                gc.DrawImageScreenChunk(_graphKatCurrent, panelGrafikkPanel, _graphFraDato, pickerDato_Graph.Value, null, _graphSelCurrent);
            else if (gc != null)
                gc.DrawImageScreenEmpty(panelGrafikkPanel, bwUpdateBigGraph.IsBusy);
        }

        private void InitGraph()
        {
            this.Update();
            pickerDato_Graph.Value = appConfig.dbTo;

            UpdateGraphFields();
            GraphPopulateSelgere();

            groupGraphChoices.Enabled = true;
        }

        private void GraphPopulateSelgere()
        {
            try
            {
                listBox_GraphSelgere.Items.Clear();

                if (selgerkodeList.Count == 0)
                {
                    UpdateSelgerkoderUI();
                }
                else
                {
                    listBox_GraphSelgere.Items.Add("ALLE");
                    listBox_GraphSelgere.Items.AddRange(selgerkodeList.ToArray());
                }
            }
            catch(Exception ex)
            {
                FormError errorMsg = new FormError("Feil oppstod under henting av selgerkoder", ex);
                errorMsg.ShowDialog(this);
            }
        }

        private void UpdateGraph()
        {
            if (EmptyDatabase() || gc == null)
                return;

            UpdateGraphFields();

            if (!bwUpdateBigGraph.IsBusy)
            {
                groupGraphChoices.Enabled = false;
                Logg.Log("Oppdaterer graf ..", Color.Black, false, true);
                bwUpdateBigGraph.RunWorkerAsync();
            }
        }

        private void bwUpdateBigGraph_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            gc.UpdateGraphChunk(_graphKatCurrent, _graphFraDato, pickerDato_Graph.Value, bwUpdateBigGraph, _graphSelCurrent);
        }

        private void bwUpdateBigGraph_Completed(object sender, AsyncCompletedEventArgs e)
        {
            ProgressStop();
            PaintGraph();
            groupGraphChoices.Enabled = true;
            if (!_graphReqStop)
                Logg.Log("Graf oppdatert.", null, false, true);
            else
            {
                Logg.Log("Graf stoppet.");
                _graphReqStop = false;
            }
        }

        private void moveDateGraph(int m = 0, bool reload = false)
        {
            if (!EmptyDatabase())
            {

                var d = pickerDato_Graph.Value;
                if (m == 1) // gå tilbake en måned
                {
                    if (appConfig.dbFrom.Date <= d.AddMonths(-1))
                        pickerDato_Graph.Value = d.AddMonths(-1);
                    else
                        pickerDato_Graph.Value = appConfig.dbFrom;
                }
                if (m == 2) // gå tilbake en dag
                {
                    if (appConfig.ignoreSunday)
                    {
                        if (appConfig.dbFrom.Date <= d.AddDays(-1) && d.AddDays(-1).DayOfWeek != DayOfWeek.Sunday)
                            pickerDato_Graph.Value = d.AddDays(-1);
                        if (appConfig.dbFrom.Date <= d.AddDays(-2) && d.AddDays(-1).DayOfWeek == DayOfWeek.Sunday)
                            pickerDato_Graph.Value = d.AddDays(-2);
                    }
                    else
                    {
                        if (appConfig.dbFrom.Date <= d.AddDays(-1))
                            pickerDato_Graph.Value = d.AddDays(-1);
                    }
                }
                if (m == 3) // gå fram en dag
                {
                    if (appConfig.ignoreSunday)
                    {
                        if (appConfig.dbTo.Date >= d.AddDays(1) && d.AddDays(1).DayOfWeek != DayOfWeek.Sunday)
                            pickerDato_Graph.Value = d.AddDays(1);
                        if (appConfig.dbTo.Date >= d.AddDays(2) && d.AddDays(1).DayOfWeek == DayOfWeek.Sunday)
                            pickerDato_Graph.Value = d.AddDays(2);
                    }
                    else
                    {
                        if (appConfig.dbTo.Date >= d.AddDays(1))
                            pickerDato_Graph.Value = d.AddDays(1);
                    }
                }
                if (m == 4) // gå fram en måned
                {
                    if (appConfig.dbTo.Date >= d.AddMonths(1))
                        pickerDato_Graph.Value = d.AddMonths(1);
                    else
                        pickerDato_Graph.Value = appConfig.dbTo;
                }
                d = pickerDato_Graph.Value;
                if (d.Date >= appConfig.dbTo.Date)
                {
                    buttonGraphF.Enabled = false; // fremover knapp
                    buttonGraphFF.Enabled = false; // fremover knapp
                }
                else
                {
                    buttonGraphF.Enabled = true; // fremover knapp
                    buttonGraphFF.Enabled = true; // fremover knapp
                }
                if (d.Date <= appConfig.dbFrom.Date)
                {
                    buttonGraphBF.Enabled = false; // bakover knapp
                    buttonGraphB.Enabled = false; // bakover knapp
                }
                else
                {
                    buttonGraphBF.Enabled = true; // bakover knapp
                    buttonGraphB.Enabled = true; // bakover knapp
                }

                if (Loaded && reload)
                {
                    UpdateGraph();
                }
            }
        }


        private void RunGraph(string argKat)
        {
            if (IsBusy())
                return;

            if (!EmptyDatabase())
            {
                groupRankingChoices.Enabled = false;
                bwGraph.RunWorkerAsync(argKat);
            }
            else
                webHTML.Navigate(htmlImport);
        }

        private void bwGraph_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            string value = (string)e.Argument;
            ViewGraph(value, false, bwGraph);
        }

        private void ReportProgressCustom(decimal current, StatusProgress status)
        {
            try
            {
                decimal total = status.total;
                decimal percent = 0;

                if (total != 0)
                    percent = current / total;
                decimal gap = status.end - status.start;

                if (gap != 0)
                {
                    decimal gapValue = gap * percent;

                    if (percent >= 0 && percent <= 100)
                        processing.SetValue = (int)(status.start + gapValue);

                    Logg.Status(status.text + " (" + (current + 1) + " av " + total + ")");
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void bwGraph_Completed(object sender, AsyncCompletedEventArgs e)
        {
            ProgressStop();
            if (!IsBusy(true))
                Logg.Status("Klar.");
            groupRankingChoices.Enabled = true;
        }


        private void SelectGraph(string katArg)
        {
            button_GraphDatamaskiner.BackColor = SystemColors.ControlLight;
            button_GraphNettbrett.BackColor = SystemColors.ControlLight;
            button_GraphTver.BackColor = SystemColors.ControlLight;
            button_GraphMobiler.BackColor = SystemColors.ControlLight;
            button_GraphButikk.BackColor = SystemColors.ControlLight;
            button_GraphTjen.BackColor = SystemColors.ControlLight;
            button_GraphKnowHow.BackColor = SystemColors.ControlLight;

            _graphKatCurrent = katArg;

            if (katArg == "Data")
                button_GraphDatamaskiner.BackColor = Color.LightSkyBlue;
            else if (katArg == "Nettbrett")
                button_GraphNettbrett.BackColor = Color.LightSkyBlue;
            else if (katArg == "AudioVideo")
                button_GraphTver.BackColor = Color.LightSkyBlue;
            else if (katArg == "Tele")
                button_GraphMobiler.BackColor = Color.LightSkyBlue;
            else if (katArg == "KnowHow")
                button_GraphKnowHow.BackColor = Color.LightSkyBlue;
            else if (katArg == "Butikk")
                button_GraphButikk.BackColor = Color.LightSkyBlue;
            else if (katArg == "Oversikt")
                button_GraphTjen.BackColor = Color.LightSkyBlue;

            this.Update();
        }


























        private Bitmap DrawToBitmapBigEmpty(Bitmap b, int argX, int argY)
        {
            using (Graphics g = Graphics.FromImage(b))
            {
                float X = argX;
                float Y = argY;

                g.Clear(Color.White);
                var bmpPicture = new Bitmap(argX, argY);
                g.DrawImage(bmpPicture, X, Y);

                var pen = new Pen(Color.Black, 1);
                var font = new Font("Helvetica", 14, FontStyle.Regular);

                g.DrawRectangle(pen, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1)));
                g.DrawString("Trykk \"Last inn\" for å starte.", font, new SolidBrush(Color.Gray), 22, 14);
            }
            return b;
        }

        public static DateTime GetFirstDayOfMonth(DateTime givenDate)
        {
            return new DateTime(givenDate.Year, givenDate.Month, 1);
        }

        public static DateTime GetLastDayOfMonth(DateTime givenDate)
        {
            return GetFirstDayOfMonth(givenDate).AddMonths(1).Subtract(new TimeSpan(1, 0, 0, 0, 0));
        }

        private static Color Tint(Color source, Color tint, decimal alpha)
        {
            //(tint -source)*alpha + source
            int red = Convert.ToInt32(((tint.R - source.R)*alpha + source.R));
            int blue = Convert.ToInt32(((tint.B - source.B)*alpha + source.B));
            int green = Convert.ToInt32(((tint.G - source.G)*alpha + source.G));
            return Color.FromArgb(255, red, green, blue);
        }

        private string getVKstring(string arg = "")
        {
            DataTable dtVk;
            if (!String.IsNullOrEmpty(arg))
                dtVk = database.GetSqlDataTable("SELECT Varekode FROM tblVarekoder WHERE Kategori = '" + arg + "' AND Synlig = 1 ORDER by Id");
            else
                dtVk = database.GetSqlDataTable("SELECT Varekode FROM tblVarekoder AND Synlig = 1 ORDER by Id");
            if (dtVk != null)
            {
                string vkSQL;
                int size = dtVk.Rows.Count;
                if (size > 0)
                {
                    vkSQL = " AND (";
                    for (int i = 0; i < size; i++)
                    {
                        vkSQL += "Varekode = '" + dtVk.Rows[i][0] + "' ";
                        if ((i + 1) < size)
                            vkSQL += " OR ";
                        else
                            vkSQL += ") ";
                    }

                    return vkSQL;
                }
                return "";
            }
            return "";
        }

        public string[] listAvd()
        {
            try
            {
                string SQL = "SELECT DISTINCT Avdeling FROM tblSalg WHERE Avdeling < 1700";
                Logg.Debug("Avdeling søk: Henter ut unike butikker..");
                DataTable dtAvd = database.GetSqlDataTable(SQL); 
                if (dtAvd != null)
                {
                    int size = dtAvd.Rows.Count;
                    string[] array = new string[size];
                    int i;
                    for (i = 0; i < size; i++)
                    {
                        array[i] = dtAvd.Rows[i][0].ToString();
                    }
                    Logg.Debug("Avdeling søk: Ferdig! Fant " + array.Length + " butikker.");
                    return array;
                }
                return null;
            }
            catch (SqlCeException ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private void PaintHistTopNew()
        {
            try
            {
                if (topgraph == null || bwUpdateTopGraph.IsBusy || !graphPanelTop.Visible)
                    return;

                int gW = graphPanelTop.Width;
                int gH = graphPanelTop.Height;
                using (var b = new Bitmap(gW, gH))
                {
                    Bitmap graphbitmap = topgraph.DrawToBitmapTop(b, gW, gH);
                    Graphics g1;
                    g1 = graphPanelTop.CreateGraphics();
                    g1.DrawImage(graphbitmap, 0, 0, graphbitmap.Width, graphbitmap.Height);
                    g1.Dispose();
                    graphbitmap.Dispose();
                }
            }
            catch
            {
            }
        }
    }


    public class StatusProgress
    {
        public int total { get; set; }
        public string text { get; set; }
        public int start { get; set; }
        public int end { get; set; }
        public StatusProgress(int totalArg, string textArg, int startArg = 0, int endArg = 100)
        {
            this.total = totalArg;
            this.text = textArg;
            this.start = startArg;
            this.end = endArg;
        }

    }
}

