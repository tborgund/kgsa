using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace KGSA
{
    public partial class GraphClass : IDisposable
    {
        FormMain main;
        public DateTime dtFra = DateTime.Now;
        public DateTime dtTil = DateTime.Now;
        private List<string> Favoritter;
        private bool disposed;

        private DataTable dataKnowHow;
        private DateTime dataKnowHowFrom;
        private DateTime dataKnowHowTo;
        private DataTable dataData;
        private DateTime dataDataFrom;
        private DateTime dataDataTo;
        private DataTable dataNettbrett;
        private DateTime dataNettbrettFrom;
        private DateTime dataNettbrettTo;
        private DataTable dataAudioVideo;
        private DateTime dataAudioVideoFrom;
        private DateTime dataAudioVideoTo;
        private DataTable dataTele;
        private DateTime dataTeleFrom;
        private DateTime dataTeleTo;
        private DataTable dataButikk;
        private DateTime dataButikkFrom;
        private DateTime dataButikkTo;
        private DataTable dataOversikt;
        private DateTime dataOversiktFrom;
        private DateTime dataOversiktTo;
        private DataTable dataSelger;
        private DateTime dataSelgerFrom;
        private DateTime dataSelgerTo;
        private string lastSelgerkode = "";
        private string lastSelgerkodeKat = "";

        public GraphClass(FormMain form, DateTime dtFraArg, DateTime dtTilArg)
        {
            this.main = form;
            dtFra = dtFraArg;
            dtTil = dtTilArg;
            if (main.appConfig.favVis)
                Favoritter = FormMain.Favoritter;
            else
                Favoritter = new List<string> { main.appConfig.Avdeling.ToString() };
        }

        public GraphClass(FormMain form)
        {
            this.main = form;
            if (main.appConfig.favVis)
                Favoritter = FormMain.Favoritter;
            else
                Favoritter = new List<string> { main.appConfig.Avdeling.ToString() };
        }

        public void UpdateGraphChunk(string argKat, DateTime dtFraArg, DateTime dtTilArg, BackgroundWorker bw, string sk = "")
        {
            Logg.Debug("Ber om oppdatering av grafikk: " + argKat + " | Fra: " + dtFraArg.ToShortDateString() + " | Til: " + dtTilArg.ToShortDateString());

            if (sk != "" && (argKat == "Data" || argKat == "Nettbrett" || argKat == "Tele" || argKat == "AudioVideo" || argKat == "KnowHow"))
            {
                if (lastSelgerkode != sk || lastSelgerkodeKat != argKat)
                {
                    if (dataSelger != null)
                        dataSelger.Clear();
                    dataSelgerFrom = FormMain.rangeMin;
                    dataSelgerTo = FormMain.rangeMin;
                    lastSelgerkode = sk;
                    lastSelgerkodeKat = argKat;
                }

                if (dataSelger != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataSelgerFrom.ToShortDateString() + " | Til: " + dataSelgerTo.ToShortDateString() + " | Count: " + dataSelger.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataSelgerFrom.Date && dtTilArg.Date <= dataSelgerTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataSelgerTo.Date || dtFraArg.Date > dataSelgerFrom.Date) && dataSelger != null && dataSelgerFrom.Date != dataSelgerTo.Date)
                    {
                        Logg.Debug("intersection: true");
                        if (dataSelgerTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataSelgerFrom.AddDays(-1), bw, lastSelgerkode);
                            dt.Merge(this.dataSelger);
                            this.dataSelger = dt;
                            dataSelgerFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataSelger.Merge(UpdateGraphInternalChunk(argKat, dataSelgerTo.AddDays(1), dtTilArg, bw, lastSelgerkode));
                            dataSelgerTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataSelger = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw, lastSelgerkode);
                        dataSelgerFrom = dtFraArg;
                        dataSelgerTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataSelgerFrom.ToShortDateString() + " | Til: " + dataSelgerTo.ToShortDateString() + " Count: " + dataSelger.Rows.Count);
                return;
            }

            if (argKat == "KnowHow")
            {
                if (dataKnowHow != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataKnowHowFrom.ToShortDateString() + " | Til: " + dataKnowHowTo.ToShortDateString() + " | Count: " + dataKnowHow.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataKnowHowFrom.Date && dtTilArg.Date <= dataKnowHowTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataKnowHowTo.Date || dtFraArg.Date > dataKnowHowFrom.Date) && dataKnowHow != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataKnowHowTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataKnowHowFrom.AddDays(-1), bw);
                            if (FormMain._graphReqStop)
                                return;
                            dt.Merge(this.dataKnowHow);
                            this.dataKnowHow = dt;
                            dataKnowHowFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataKnowHow.Merge(UpdateGraphInternalChunk(argKat, dataKnowHowTo.AddDays(1), dtTilArg, bw));
                            if (FormMain._graphReqStop)
                                return;
                            dataKnowHowTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataKnowHow = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw);
                        if (FormMain._graphReqStop)
                            return;
                        dataKnowHowFrom = dtFraArg;
                        dataKnowHowTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataKnowHowFrom.ToShortDateString() + " | Til: " + dataKnowHowTo.ToShortDateString() + " Count: " + dataKnowHow.Rows.Count);
                return;
            }
            else if (argKat == "Data")
            {
                if (dataData != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataDataFrom.ToShortDateString() + " | Til: " + dataDataTo.ToShortDateString() + " | Count: " + dataData.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataDataFrom.Date && dtTilArg.Date <= dataDataTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataDataTo.Date || dtFraArg.Date > dataDataFrom.Date) && dataData != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataDataTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataDataFrom.AddDays(-1), bw);
                            if (FormMain._graphReqStop)
                                return;
                            dt.Merge(this.dataData);
                            this.dataData = dt;
                            dataDataFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataData.Merge(UpdateGraphInternalChunk(argKat, dataDataTo.AddDays(1), dtTilArg, bw));
                            if (FormMain._graphReqStop)
                                return;
                            dataDataTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataData = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw);
                        if (FormMain._graphReqStop)
                            return;
                        dataDataFrom = dtFraArg;
                        dataDataTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataDataFrom.ToShortDateString() + " | Til: " + dataDataTo.ToShortDateString() + " Count: " + dataData.Rows.Count);
                return;
            }
            else if (argKat == "Nettbrett")
            {
                if (dataNettbrett != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataNettbrettFrom.ToShortDateString() + " | Til: " + dataNettbrettTo.ToShortDateString() + " | Count: " + dataNettbrett.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataNettbrettFrom.Date && dtTilArg.Date <= dataNettbrettTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataNettbrettTo.Date || dtFraArg.Date > dataNettbrettFrom.Date) && dataNettbrett != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataNettbrettTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataNettbrettFrom.AddDays(-1), bw);
                            dt.Merge(this.dataNettbrett);
                            this.dataNettbrett = dt;
                            dataNettbrettFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataNettbrett.Merge(UpdateGraphInternalChunk(argKat, dataNettbrettTo.AddDays(1), dtTilArg, bw));
                            dataNettbrettTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataNettbrett = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw);
                        dataNettbrettFrom = dtFraArg;
                        dataNettbrettTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataNettbrettFrom.ToShortDateString() + " | Til: " + dataNettbrettTo.ToShortDateString() + " Count: " + dataNettbrett.Rows.Count);
                return;
            }
            else if (argKat == "AudioVideo")
            {
                if (dataAudioVideo != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataAudioVideoFrom.ToShortDateString() + " | Til: " + dataAudioVideoTo.ToShortDateString() + " | Count: " + dataAudioVideo.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataAudioVideoFrom.Date && dtTilArg.Date <= dataAudioVideoTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataAudioVideoTo.Date || dtFraArg.Date > dataAudioVideoFrom.Date) && dataAudioVideo != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataAudioVideoTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataAudioVideoFrom.AddDays(-1), bw);
                            dt.Merge(this.dataAudioVideo);
                            this.dataAudioVideo = dt;
                            dataAudioVideoFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataAudioVideo.Merge(UpdateGraphInternalChunk(argKat, dataAudioVideoTo.AddDays(1), dtTilArg, bw));
                            dataAudioVideoTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataAudioVideo = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw);
                        dataAudioVideoFrom = dtFraArg;
                        dataAudioVideoTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataAudioVideoFrom.ToShortDateString() + " | Til: " + dataAudioVideoTo.ToShortDateString() + " Count: " + dataAudioVideo.Rows.Count);
                return;
            }
            else if (argKat == "Tele")
            {
                if (dataTele != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataTeleFrom.ToShortDateString() + " | Til: " + dataTeleTo.ToShortDateString() + " | Count: " + dataTele.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataTeleFrom.Date && dtTilArg.Date <= dataTeleTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataTeleTo.Date || dtFraArg.Date > dataTeleFrom.Date) && dataTele != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataTeleTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateGraphInternalChunk(argKat, dtFraArg, dataTeleFrom.AddDays(-1), bw);
                            dt.Merge(this.dataTele);
                            this.dataTele = dt;
                            dataTeleFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataTele.Merge(UpdateGraphInternalChunk(argKat, dataTeleTo.AddDays(1), dtTilArg, bw));
                            dataTeleTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataTele = UpdateGraphInternalChunk(argKat, dtFraArg, dtTilArg, bw);
                        dataTeleFrom = dtFraArg;
                        dataTeleTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataTeleFrom.ToShortDateString() + " | Til: " + dataTeleTo.ToShortDateString() + " Count: " + dataTele.Rows.Count);
                return;
            }
            else if (argKat == "Oversikt")
            {
                if (dataOversikt != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataOversiktFrom.ToShortDateString() + " | Til: " + dataOversiktTo.ToShortDateString() + " | Count: " + dataOversikt.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataOversiktFrom.Date && dtTilArg.Date <= dataOversiktTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataOversiktTo.Date || dtFraArg.Date > dataOversiktFrom.Date) && dataOversikt != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataOversiktTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateOversiktGraphChunk(dtFraArg, dataOversiktFrom.AddDays(-1), bw);
                            dt.Merge(this.dataOversikt);
                            this.dataOversikt = dt;
                            dataOversiktFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataOversikt.Merge(UpdateOversiktGraphChunk(dataOversiktTo.AddDays(1), dtTilArg, bw));
                            dataOversiktTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataOversikt = UpdateOversiktGraphChunk(dtFraArg, dtTilArg, bw);
                        dataOversiktFrom = dtFraArg;
                        dataOversiktTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataOversiktFrom.ToShortDateString() + " | Til: " + dataOversiktTo.ToShortDateString() + " Count: " + dataOversikt.Rows.Count);
                return;
            }
            else if (argKat == "Butikk")
            {
                if (dataButikk != null)
                    Logg.Debug("Datoer vi har - Fra: " + dataButikkFrom.ToShortDateString() + " | Til: " + dataButikkTo.ToShortDateString() + " | Count: " + dataButikk.Rows.Count);
                else
                    Logg.Debug("Databasen var tom.");

                if (dtFraArg.Date >= dataButikkFrom.Date && dtTilArg.Date <= dataButikkTo.Date)
                {
                    Logg.Debug("intersection: all inside!");
                    return;
                }
                else
                {
                    if ((dtTilArg.Date < dataButikkTo.Date || dtFraArg.Date > dataButikkFrom.Date) && dataButikk != null)
                    {
                        Logg.Debug("intersection: true");
                        if (dataButikkTo.Date > dtTilArg.Date)
                        {
                            Logg.Debug("intersection: right");
                            DataTable dt = UpdateButikkGraphChunk(dtFraArg, dataButikkFrom.AddDays(-1), bw);
                            dt.Merge(this.dataButikk);
                            this.dataButikk = dt;
                            dataButikkFrom = dtFraArg;
                        }
                        else
                        {
                            Logg.Debug("intersection: left");
                            dataButikk.Merge(UpdateButikkGraphChunk(dataButikkTo.AddDays(1), dtTilArg, bw));
                            dataButikkTo = dtTilArg;
                        }
                    }
                    else
                    {
                        Logg.Debug("clean update!");
                        this.dataButikk = UpdateButikkGraphChunk(dtFraArg, dtTilArg, bw);
                        dataButikkFrom = dtFraArg;
                        dataButikkTo = dtTilArg;
                    }
                }
                Logg.Debug("Ny oppdatert database - Fra: " + dataButikkFrom.ToShortDateString() + " | Til: " + dataButikkTo.ToShortDateString() + " Count: " + dataButikk.Rows.Count);
                return;
            }
        }

        public void SaveImageChunk(string argKat, string argFilename, int argX, int argY, string argAgr, DateTime dtFraArg, DateTime dtTilArg, bool noUpdate = false, BackgroundWorker bw = null)
        {
            if (FormMain.stopRanking)
                return;

            Bitmap graphBitmap = DrawImageChunk(argKat, argX, argY, argAgr, dtFraArg, dtTilArg, noUpdate, bw, false);
            if (graphBitmap != null)
            {
                graphBitmap.Save(argFilename, ImageFormat.Png);
                graphBitmap.Dispose();
            }
            
        }

        public List<string> SaveGraphData()
        {
            try
            {
                List<string> doc = new List<string>() { };

                if (FormMain.stopRanking)
                    return doc;

                if (dataButikk != null)
                {
                    string data = "[ [0, 0] ";
                    for (int i = 0; i < dataButikk.Rows.Count; i++)
                    {
                        //data += ", [ " + Convert.ToDateTime(dataButikk.Rows[i]["Dato"]).Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds + ", " + dataButikk.Rows[i]["Index"] + " ]";
                        var store = (StorageButikk)dataButikk.Rows[i][2];
                        data += ", [ " + GetJavascriptTimestamp(Convert.ToDateTime(dataButikk.Rows[i][1])) + ", " + store.btokr + " ]";

                    }
                    data += " ]";
                    doc.Add(data);
                     
                }


                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Feil oppstod under lagring av graf data." };
            }
        }

        public static long GetJavascriptTimestamp(System.DateTime input)
        {
            System.TimeSpan span = new System.TimeSpan(System.DateTime.Parse("1/1/1970").Ticks);
            System.DateTime time = input.Subtract(span);
            return (long)(time.Ticks / 10000);
        }

        public void DrawImageScreenEmpty(Panel panel, bool loading)
        {
            try
            {
                if (panel != null && panel.Visible == true)
                {
                    int pWidth = panel.Width;
                    int pHeight = panel.Height;
                    using (var b = new Bitmap(pWidth, pHeight))
                    {
                        Bitmap graphBitmap;
                        graphBitmap = DrawImageScreenEmpty(pWidth, pHeight, loading);
                        if (graphBitmap != null)
                        {
                            Graphics g1;
                            g1 = panel.CreateGraphics();
                            g1.DrawImage(graphBitmap, 0, 0, graphBitmap.Width, graphBitmap.Height);
                            g1.Dispose();
                            graphBitmap.Dispose();
                        }
                        else
                        {
                            Logg.Log("Ingenting å tegne..");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private Bitmap DrawImageScreenEmpty(int width, int height, bool loading)
        {
            Bitmap b = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(b))
            {
                Bitmap myBitmap;

                if (main.appConfig.chainElkjop)
                    myBitmap = new Bitmap(KGSA.Properties.Resources.Elkjoplogo_Bg);
                else
                    myBitmap = new Bitmap(KGSA.Properties.Resources.Lefdal_Bg);
                g.Clear(Color.White);
                g.DrawImage(myBitmap, (width - 758) / 2, (height - 289) / 2);
                g.DrawRectangle(new Pen(Color.Black, 1), new Rectangle(new Point(0, 0), new Size(width - 1, height - 1))); // ramme
                if (loading)
                    g.DrawString("Laster..", new Font("Verdana", 20, FontStyle.Bold), new SolidBrush(Color.Gray), 10, 10);
            }
            return b;
        }

        public void DrawImageScreenChunk(string argKat, Panel panel, DateTime dtFraArg, DateTime dtTilArg, BackgroundWorker bw = null, string sk = "")
        {
            try
            {
                if (panel != null)
                {
                    int pWidth = panel.Width;
                    int pHeight = panel.Height;
                    if (pHeight < 100 || pWidth < 100 || panel.Visible == false)
                        return;
                    using (var b = new Bitmap(pWidth, pHeight))
                    {
                        Bitmap graphBitmap;
                        graphBitmap = DrawImageChunk(argKat, pWidth, pHeight, "", dtFraArg, dtTilArg, true, bw, true, sk);
                        if (graphBitmap != null)
                        {
                            Graphics g1;
                            g1 = panel.CreateGraphics();
                            g1.DrawImage(graphBitmap, 0, 0, graphBitmap.Width, graphBitmap.Height);
                            g1.Dispose();
                            graphBitmap.Dispose();
                        }
                        else
                        {
                            Bitmap graphBitmap2;
                            graphBitmap2 = DrawImageScreenEmpty(pWidth, pHeight, false);
                            if (graphBitmap2 != null)
                            {
                                Graphics g2;
                                g2 = panel.CreateGraphics();
                                g2.DrawImage(graphBitmap2, 0, 0, graphBitmap2.Width, graphBitmap2.Height);
                                g2.Dispose();
                                graphBitmap2.Dispose();
                            }
                            else
                            {
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Logg.Debug("Feil oppstod i DrawImageScreenChunk.", ex);
            }
        }

        public Bitmap DrawImageChunk(string argKat, int argX, int argY, string argAgr, DateTime dtFraArg, DateTime dtTilArg, bool noUpdate, BackgroundWorker bw, bool screen, string sk = "")
        {
            try
            {
                string filter = string.Format(CultureInfo.InvariantCulture, "Dato >= '{0}' AND Dato <= '{1}'", dtFraArg.ToString("o", CultureInfo.InvariantCulture), dtTilArg.ToString("o", CultureInfo.InvariantCulture));

                if (sk != "" && (argKat == "Data" || argKat == "Nettbrett" || argKat == "Tele" || argKat == "AudioVideo" || argKat == "KnowHow"))
                {
                    if (!noUpdate && dataSelger == null)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw, sk);
                    if (FormMain._graphReqStop || dataSelger == null)
                        return null;
                    DataView dv = new DataView(dataSelger);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "Salg av " + sk, argKat, argAgr, max, maxWeek, maxMonth, percent, screen, sk);
                }

                if (argKat == "KnowHow")
                {
                    if (dataKnowHow == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataKnowHow == null)
                        return null;
                    DataView dv = new DataView(dataKnowHow);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "Produkter", "KnowHow", argAgr, max, maxWeek, maxMonth, percent, screen);
                }
                else if (argKat == "Data")
                {
                    if (dataData == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataData == null)
                        return null;
                    DataView dv = new DataView(dataData);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "Datamaskiner", "Data", argAgr, max, maxWeek, maxMonth, percent, screen);
                }
                else if (argKat == "Nettbrett")
                {
                    if (dataNettbrett == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataNettbrett == null)
                        return null;
                    DataView dv = new DataView(dataNettbrett);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "Nettbrett", "Data", argAgr, max, maxWeek, maxMonth, percent, screen);
                }
                else if (argKat == "AudioVideo")
                {
                    if (dataAudioVideo == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataAudioVideo == null)
                        return null;
                    DataView dv = new DataView(dataAudioVideo);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "TVer", "Lyd og Bilde", argAgr, max, maxWeek, maxMonth, percent, screen);
                }
                else if (argKat == "Tele")
                {
                    if (dataTele == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataTele == null)
                        return null;
                    DataView dv = new DataView(dataTele);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float max = 0f, sumProduct = 0, sumService = 0, percent = 0, sumProductMonth = 0, maxMonth = 0, sumProductWeek = 0, maxWeek = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageData)dt.Rows[i][2];
                        sumProduct += store.product;
                        sumService += store.service;

                        float per = CalcPercent(sumService, sumProduct);
                        if (store.product > max)
                            max = store.product;
                        if (percent < per)
                            percent = per;

                        sumProductMonth += store.product;
                        sumProductWeek += store.product;
                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumProductMonth)
                                maxMonth = sumProductMonth;
                            sumProductMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumProductWeek)
                                maxWeek = sumProductWeek;
                            sumProductWeek = 0;
                        }
                    }

                    return DrawToBitmapChunk(argX, argY, dt, "Mobiler", "Tele", argAgr, max, maxWeek, maxMonth, percent, screen);
                }
                else if (argKat == "Oversikt")
                {
                    if (dataOversikt == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataOversikt == null)
                        return null;
                    DataView dv = new DataView(dataOversikt);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    float maxWeek = 0, maxMonth = 0, sumTotal = 0, max = 0, sumWeek = 0, sumMonth = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageButikk)dt.Rows[i][2];
                        sumTotal += store.btokr;
                        sumWeek += store.btokr;
                        sumMonth += store.btokr;

                        if (store.btokr > max)
                            max = store.btokr;

                        if (Convert.ToDateTime(dt.Rows[i][1]).Day == 1)
                        {
                            if (maxMonth < sumMonth)
                                maxMonth = sumMonth;
                            sumMonth = 0;
                        }
                        if (Convert.ToDateTime(dt.Rows[i][1]).DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumWeek)
                                maxWeek = sumWeek;
                            sumWeek = 0;
                        }
                    }

                    return DrawToBitmapOversiktChunk(argX, argY, dt, "Oversikt", "Tjenester", argAgr, max, maxWeek, maxMonth, 1, screen);
                }
                else if (argKat == "Butikk")
                {
                    if (dataButikk == null || !noUpdate)
                        UpdateGraphChunk(argKat, dtFraArg, dtTilArg, bw);

                    if (FormMain._graphReqStop || dataButikk == null)
                        return null;
                    DataView dv = new DataView(dataButikk);
                    dv.RowFilter = filter;
                    DataTable dt = dv.ToTable();
                    int count = dt.Rows.Count;
                    if (count == 0)
                        return null;
                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[count - 1][1];
                    float maxWeek = 0, maxMonth = 0, sumTotal = 0, max = 0, sumWeek = 0, sumMonth = 0;
                    for (int i = 0; i < count; i++)
                    {
                        var store = (StorageButikk)dt.Rows[i][2];
                        var date = Convert.ToDateTime(dt.Rows[i][1]);
                        sumTotal += store.btokr;
                        sumWeek += store.btokr;
                        sumMonth += store.btokr;

                        if (store.btokr > max)
                            max = store.btokr;

                        if (date.Day == 1)
                        {
                            if (maxMonth < sumMonth)
                                maxMonth = sumMonth;
                            sumMonth = 0;
                        }
                        if (date.DayOfWeek == DayOfWeek.Monday)
                        {
                            if (maxWeek < sumWeek)
                                maxWeek = sumWeek;
                            sumWeek = 0;
                        }
                        if (date.AddDays(1).Day == 1 && !(firstdate.Month == 5 && firstdate.Day == 1) && !FormMain.datoPeriodeVelger && dt.Rows.Count != i + 1 && !(!main.appConfig.graphHitrateMTD && screen))
                            sumTotal = 0;
                    }

                    return DrawToBitmapButikkChunk(argX, argY, dt, "Butikk", "Butikk", argAgr, max, maxWeek, maxMonth, 1, screen, sumTotal);
                }

                return null;
            }
            catch (IndexOutOfRangeException)
            {
                return null;
            }

            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private DataTable UpdateGraphInternalChunk(string argKat, DateTime dtFraArg, DateTime dtTilArg, BackgroundWorker bw, string sk = "")
        {
            Logg.Debug("Oppdaterer graf: " + argKat + " | Fra: " + dtFraArg.ToShortDateString() + " | Til: " + dtTilArg.ToShortDateString());

            var dt = new DataTable();
            string skString = "";
            if (sk != "")
                skString = " AND Selgerkode = '" + sk + "' ";

            string product = "";
            if (argKat == "KnowHow")
                product = "Varegruppe = '531' OR Varegruppe = '533' OR Varegruppe = '534' OR Varegruppe = '224' OR Varegruppe = '431'";
            else if (argKat == "Data")
                product = "Varegruppe = '531' OR Varegruppe = '533'";
            else if (argKat == "Nettbrett")
                product = "Varegruppe = '534'";
            else if (argKat == "AudioVideo")
                product = "Varegruppe = '224'";
            else if (argKat == "Tele")
                product = "Varegruppe = '431'";
            else
                return dt;

            dt.Columns.Add("Index", typeof(int));
            dt.Columns.Add("Dato", typeof(DateTime));

            DataTable dtResult;
            int days = (dtTilArg - dtFraArg).Days + 1;

            for (int d = 0; d < Favoritter.Count; d++)
            {
                if (sk != "" && d > 0)
                    break;

                string command = "SELECT SUM(CASE WHEN " + product + " THEN Antall ELSE 0 END) AS Product, "
                    + "SUM(CASE WHEN " + GetVarekodeSqlString(argKat) + " THEN Antall ELSE 0 END) AS Service, "
                    + "Dato FROM tblSalg WHERE (Avdeling = '" + Favoritter[d] + "') AND (Dato >= '"
                    + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd")
                    + "') " + skString + " GROUP BY Dato";

                dtResult = main.database.GetSqlDataTable(command);

                dt.Columns.Add(Favoritter[d], typeof(StorageData));
                if (dt.Rows.Count == 0)
                {
                    for (int o = 0; o < days; o++)
                    {
                        if (bw != null)
                            if (bw.WorkerReportsProgress)
                                bw.ReportProgress(o, new StatusProgress(days, "Oppdaterer graf..", 0, 40));
                        if (FormMain._graphReqStop)
                            return null;

                        DataRow row = dt.NewRow();
                        row[0] = o; // Index
                        int count = dtResult.Rows.Count;
                        for (int i = 0; i < count; i++)
                        {
                            if (Convert.ToDateTime(dtResult.Rows[i][2]) == dtFraArg.AddDays(o))
                            {
                                row[1] = (DateTime)dtResult.Rows[i][2]; // Dato
                                var store = new StorageData((int)dtResult.Rows[i][0], (int)dtResult.Rows[i][1]);
                                row[d + 2] = store;
                                break;
                            }
                            else
                            {
                                row[1] = dtFraArg.AddDays(o); // Dato
                                row[d + 2] = new StorageData(0, 0);
                            }
                        }
                        dt.Rows.Add(row);
                    }
                }
                else
                    for (int b = 0; b < dt.Rows.Count; b++)
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                            if ((DateTime)dt.Rows[b][1] == (DateTime)dtResult.Rows[i][2])
                                dt.Rows[b][Favoritter[d]] = new StorageData((int)dtResult.Rows[i][0], (int)dtResult.Rows[i][1]);

                for (int b = 0; b < dt.Rows.Count; b++)
                    if (DBNull.Value.Equals(dt.Rows[b][Favoritter[d]]))
                        dt.Rows[b][Favoritter[d]] = new StorageData(0, 0);
            }

            if (main.appConfig.graphAdvanced && sk == "")
            {
                if (FormMain._graphReqStop)
                    return null;

                dt.Columns.Add("TOP", typeof(StorageTop));

                string command = "SELECT Dato, Selgerkode, SUM(CASE WHEN " + GetVarekodeSqlString(argKat) + " THEN Antall ELSE 0 END) AS Antall FROM tblSalg WHERE (Avdeling = '" + Favoritter[0] + "') AND (Dato >= '" + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd") + "') GROUP BY Dato, Selgerkode";
                dtResult = main.database.GetSqlDataTable(command);

                if (dtResult.Rows.Count == 0)
                    return dt;

                int count = dt.Rows.Count;
                for (int b = 0; b < count; b++)
                {
                    if (bw != null)
                        if (bw.WorkerReportsProgress)
                            bw.ReportProgress(b, new StatusProgress(count, "Oppdaterer graf detaljer.. ", 40, 100));
                    if (FormMain._graphReqStop)
                        return null;

                    DateTime date = (DateTime)dt.Rows[b][1];

                    DataView view = new DataView(dtResult);
                    view.RowFilter = string.Format(CultureInfo.InvariantCulture, "Dato = '{0}'", date.ToString("o", CultureInfo.InvariantCulture));
                    view.Sort = "Antall DESC";
                    DataTable res = view.ToTable();

                    var store = new StorageTop();

                    int resCount = res.Rows.Count;
                    for (int i = 0; i < resCount && i < 5; i++)
                    {
                        int value = Convert.ToInt32(res.Rows[i][2]);
                        if (value > 0)
                        {
                            store.selgere.Add(new StorageSelger(res.Rows[i][1].ToString(), value));
                            dt.Rows[b]["TOP"] = store;
                        }
                    }
                }
            }

            return dt;
        }

        private DataTable UpdateOversiktGraphChunk(DateTime dtFraArg, DateTime dtTilArg, BackgroundWorker bw)
        {
            Logg.Debug("Oppdaterer graf: Oversikt | Fra: " + dtFraArg.ToShortDateString() + " | Til: " + dtTilArg.ToShortDateString());

            var dt = new DataTable();

            dt.Columns.Add("Index", typeof(int));
            dt.Columns.Add("Dato", typeof(DateTime));

            DataTable dtResult;
            int days = (dtTilArg - dtFraArg).Days + 1;

            for (int d = 0; d < Favoritter.Count; d++)
            {
                string sqlstring = "SELECT SUM(Btokr) AS Btokr, SUM(Salgspris) AS Salgspris, SUM(Salgspris / Mva) AS Salgsexmva, Dato FROM tblSalg " +
                    "WHERE (Avdeling = '" + main.appConfig.Avdeling + "') AND (Dato >= '" + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd") + "') GROUP BY Dato";

                dtResult = main.database.GetSqlDataTable(sqlstring);

                dt.Columns.Add(Favoritter[d], typeof(StorageButikk));
                if (dt.Rows.Count == 0)
                {
                    for (int o = 0; o < days; o++)
                    {
                        if (bw != null)
                            if (bw.WorkerReportsProgress)
                                bw.ReportProgress(o, new StatusProgress(days, "Oppdaterer graf.. ", 0, 40));
                        if (FormMain._graphReqStop)
                            return null;

                        DataRow row = dt.NewRow();
                        row[0] = o; // Index
                        int count = dtResult.Rows.Count;
                        for (int i = 0; i < count; i++)
                        {
                            if (Convert.ToDateTime(dtResult.Rows[i][3]) == dtFraArg.AddDays(o))
                            {
                                row[1] = (DateTime)dtResult.Rows[i][3]; // Dato
                                var store = new StorageButikk(Convert.ToInt32(dtResult.Rows[i][0]), Convert.ToInt32(dtResult.Rows[i][1]), Convert.ToInt32(dtResult.Rows[i][2]));
                                row[d + 2] = store;
                                break;
                            }
                            else
                            {
                                row[1] = dtFraArg.AddDays(o); // Dato
                                row[d + 2] = new StorageButikk(0, 0, 0);
                            }
                        }
                        dt.Rows.Add(row);
                    }
                }
                else
                    for (int b = 0; b < dt.Rows.Count; b++)
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                            if ((DateTime)dt.Rows[b][1] == (DateTime)dtResult.Rows[i][3])
                                dt.Rows[b][Favoritter[d]] = new StorageButikk(Convert.ToInt32(dtResult.Rows[i][0]), Convert.ToInt32(dtResult.Rows[i][1]), Convert.ToInt32(dtResult.Rows[i][2]));

                for (int b = 0; b < dt.Rows.Count; b++)
                    if (DBNull.Value.Equals(dt.Rows[b][Favoritter[d]]))
                        dt.Rows[b][Favoritter[d]] = new StorageButikk(0, 0, 0);
            }


            dt.Columns.Add("Tjenester", typeof(StorageTjenester));

            Logg.Status("Oppdaterer graf detaljer..");

            string command = "SELECT Dato, " +
                "SUM(CASE WHEN Varegruppe LIKE '961%' THEN Btokr ELSE 0 END) AS Finans, " +
                "SUM(CASE WHEN Varegruppe LIKE '_83' AND Varekode LIKE 'MOD%' THEN Btokr ELSE 0 END) AS Ta, " +
                "SUM(CASE WHEN Varekode LIKE 'ELSTROM%' OR Varekode LIKE 'ELRABATT%' THEN Btokr ELSE 0 END) AS Strom, " +
                "SUM(CASE WHEN " + GetVarekodeSqlString() + " THEN Btokr ELSE 0 END) AS Tjen, " +
                "SUM(CASE WHEN Varegruppe LIKE '_83' AND Varekode LIKE 'MOD%' THEN Salgspris ELSE 0 END) AS TaOmset " +
                "FROM tblSalg WHERE (Avdeling = '" + Favoritter[0] + "') AND (Dato >= '" + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd") + "') GROUP BY Dato";
            dtResult = main.database.GetSqlDataTable(command);

            int countDetaljer = dt.Rows.Count;
            for (int b = 0; b < countDetaljer; b++)
            {
                if (bw != null)
                    if (bw.WorkerReportsProgress)
                        bw.ReportProgress(b, new StatusProgress(countDetaljer, "Oppdaterer graf detaljer.. ", 40, 100));
                if (FormMain._graphReqStop)
                    return null;

                DateTime date = (DateTime)dt.Rows[b][1];

                DataView view = new DataView(dtResult);
                view.RowFilter = string.Format(CultureInfo.InvariantCulture, "Dato = '{0}'", date.ToString("o", CultureInfo.InvariantCulture));
                DataTable day = view.ToTable();

                if (day.Rows.Count > 0)
                    dt.Rows[b]["Tjenester"] = new StorageTjenester(Convert.ToInt32(day.Rows[0][1]), Convert.ToInt32(day.Rows[0][2]), Convert.ToInt32(day.Rows[0][3]), Convert.ToInt32(day.Rows[0][4]), Convert.ToInt32(day.Rows[0][5]));
                else
                    dt.Rows[b]["Tjenester"] = new StorageTjenester(0, 0, 0, 0, 0);
            }


            return dt;
        }

        private DataTable UpdateButikkGraphChunk(DateTime dtFraArg, DateTime dtTilArg, BackgroundWorker bw)
        {
            Logg.Debug("Oppdaterer graf: Butikk | Fra: " + dtFraArg.ToShortDateString() + " | Til: " + dtTilArg.ToShortDateString());

            var dt = new DataTable();
            dt.Columns.Add("Index", typeof(int));
            dt.Columns.Add("Dato", typeof(DateTime));

            DataTable dtResult;
            int days = (dtTilArg - dtFraArg).Days + 1;

            for (int d = 0; d < Favoritter.Count; d++)
            {
                string sqlstring = "SELECT SUM(Btokr) AS Btokr, SUM(Salgspris) AS Salgspris, SUM(Salgspris / Mva) AS Salgsexmva, Dato FROM tblSalg " +
                    "WHERE (Avdeling = '" + Favoritter[d] + "') AND (Dato >= '" + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd") + "') GROUP BY Dato";
                dtResult = main.database.GetSqlDataTable(sqlstring);

                dt.Columns.Add(Favoritter[d], typeof(StorageButikk));
                if (dt.Rows.Count == 0)
                {
                    for (int o = 0; o < days; o++)
                    {
                        if (bw != null)
                            if (bw.WorkerReportsProgress)
                                bw.ReportProgress(o, new StatusProgress(days, "Oppdaterer graf.. ", 0, 40));
                        if (FormMain._graphReqStop)
                            return null;

                        DataRow row = dt.NewRow();
                        row[0] = o; // Index
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                        {
                            if (Convert.ToDateTime(dtResult.Rows[i][3]) == dtFraArg.AddDays(o))
                            {
                                row[1] = (DateTime)dtResult.Rows[i][3]; // Dato
                                var store = new StorageButikk(Convert.ToInt32(dtResult.Rows[i][0]), Convert.ToInt32(dtResult.Rows[i][1]), Convert.ToInt32(dtResult.Rows[i][2]));
                                row[d + 2] = store;
                                break;
                            }
                            else
                            {
                                row[1] = dtFraArg.AddDays(o); // Dato
                                row[d + 2] = new StorageButikk(0, 0, 0);
                            }
                        }
                        dt.Rows.Add(row);
                    }
                }
                else
                    for (int b = 0; b < dt.Rows.Count; b++)
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                            if ((DateTime)dt.Rows[b][1] == (DateTime)dtResult.Rows[i][3])
                                dt.Rows[b][Favoritter[d]] = new StorageButikk(Convert.ToInt32(dtResult.Rows[i][0]), Convert.ToInt32(dtResult.Rows[i][1]), Convert.ToInt32(dtResult.Rows[i][2]));

                for (int b = 0; b < dt.Rows.Count; b++)
                    if (DBNull.Value.Equals(dt.Rows[b][Favoritter[d]]))
                        dt.Rows[b][Favoritter[d]] = new StorageButikk(0, 0, 0);
            }

            if (main.appConfig.rankingCompareLastyear > 0 && dt.Columns.Count > 2)
            {
                string sqlstring = "SELECT SUM(Btokr) AS Btokr, Dato FROM tblSalg " +
                    "WHERE (Avdeling = '" + main.appConfig.Avdeling + "') AND (Dato >= '" + dtFraArg.AddYears(-1).ToString("yyy-MM-dd") +
                    "') AND (Dato <= '" + dtTilArg.AddYears(-1).ToString("yyy-MM-dd") + "') GROUP BY Dato";

                using (dtResult = main.database.GetSqlDataTable(sqlstring))
                {
                    for (int b = 0; b < dt.Rows.Count; b++)
                        for (int i = 0; i < dtResult.Rows.Count; i++)
                            if (Convert.ToDateTime(dt.Rows[b][1]) == Convert.ToDateTime(dtResult.Rows[i][1]).AddYears(1))
                            {
                                var store = (StorageButikk)dt.Rows[b][2];
                                store.ifjor_btokr = Convert.ToInt32(dtResult.Rows[i][0]);
                                dt.Rows[b][2] = store;
                            }
                }
            }

            if (main.appConfig.graphAdvanced)
            {
                Logg.Status("Oppdaterer graf detaljer..");

                dt.Columns.Add("TOP", typeof(StorageTop));

                string command = "SELECT Dato, Selgerkode, SUM(Btokr) AS Btokr FROM tblSalg WHERE (Avdeling = '" + Favoritter[0] + "') AND (Dato >= '" + dtFraArg.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dtTilArg.ToString("yyy-MM-dd") + "') GROUP BY Dato, Selgerkode";
                dtResult = main.database.GetSqlDataTable(command);

                int count = dt.Rows.Count;
                for (int b = 0; b < dt.Rows.Count; b++)
                {
                    if (bw != null)
                        if (bw.WorkerReportsProgress)
                            bw.ReportProgress(b, new StatusProgress(count, "Oppdaterer graf detaljer.. ", 40, 100));
                    if (FormMain._graphReqStop)
                        return null;

                    DateTime date = (DateTime)dt.Rows[b][1];

                    DataView view = new DataView(dtResult);
                    view.RowFilter = string.Format(CultureInfo.InvariantCulture, "Dato = '{0}'", date.ToString("o", CultureInfo.InvariantCulture));
                    view.Sort = "Btokr DESC";
                    DataTable res = view.ToTable();

                    var store = new StorageTop();

                    for (int i = 0; i < res.Rows.Count && i < 6; i++)
                    {
                        int value = Convert.ToInt32(res.Rows[i][2]);
                        if (value > 0)
                        {
                            store.selgere.Add(new StorageSelger(res.Rows[i][1].ToString(), value));
                            dt.Rows[b]["TOP"] = store;
                        }

                    }
                }
            }

            return dt;
        }

        private Bitmap DrawToBitmapChunk(int argX, int argY, DataTable dt, string argTitle, string argCaption, string argAgr, float argMax = 15, float argMaxWeek = 40, float argMaxMonth = 100, float argPercent = 1, bool screen = false, string sk = "")
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;
                    if (screen)
                        dpi = main.appConfig.graphScreenDPI;
                    int fontHeight = Convert.ToInt32(29 * dpi);
                    int boxLength = Convert.ToInt32(22 * dpi);
                    int fontSepHeight = Convert.ToInt32((fontHeight / 5) * dpi);

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (dt.Rows.Count == 0)
                    {
                        g.DrawString("Mangler transaksjoner!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];
                    int offsetY = Convert.ToInt32(40 * dpi);
                    int offsetX = 150;
                    float X = argX - offsetX;
                    float Y = argY - offsetY;
                    int pSize = dt.Rows.Count;
                    float Hstep = X / pSize;
                    
                    if (argAgr == "")
                    {
                        if (Hstep > (40 * dpi))
                            argAgr = "dag";
                        else if (Hstep <= (40 * dpi) && Hstep >= (8 * dpi))
                            argAgr = "uke";
                        else if (Hstep < (8 * dpi))
                            argAgr = "måned";
                    }

                    float Vstep = Y / (argMax * 1.1f);
                    if (argAgr == "uke")
                        Vstep = Y / (argMaxWeek * 1.1f);
                    if (argAgr == "måned")
                        Vstep = Y / (argMaxMonth * 1.1f);

                    float zoom = 1;
                    if (!screen || main.appConfig.graphScreenZoom && screen)
                    {
                        if (argPercent < 0.40F)
                            zoom = 2;
                        if (argPercent < 0.20F)
                            zoom = 3;
                        if (argPercent < 0.15F)
                            zoom = 4;
                        if (argPercent < 0.10F)
                            zoom = 6;
                    }

                    PointF p1, p2, p3, p4, p6;
                    Pen bPen = new Pen(Color.LightGray, 3 * dpi);
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen1 = new Pen(Color.Black, 3 * dpi);
                    Pen pen2 = new Pen(Color.Green, 6 * dpi);
                    Font fontBig = new Font("Helvetica", 20 * dpi, FontStyle.Bold);
                    Font fontNormal = new Font("Helvetica", 18 * dpi, FontStyle.Regular);
                    Color tjenColor = System.Drawing.ColorTranslator.FromHtml("#80c34a");
                    List<string> toppselgere = new List<string> { };
                    DagSjekk agr = new DagSjekk(argAgr);
                    List<string> fav = new List<string> { };
                    if (sk != "")
                        fav.Add(main.appConfig.Avdeling.ToString());
                    else
                        fav = Favoritter;

                    g.DrawString(main.avdeling.Get(main.appConfig.Avdeling) + " " + argCaption,
                        new Font("Verdana", 36 * dpi, FontStyle.Bold), new SolidBrush(Color.LightGray), 400, 0);

                    var step = Y / 4;
                    g.DrawLine(bPen, new Point(0, (int)step * 3), new Point((int)X, (int)step * 3));
                    g.DrawLine(bPen, new Point(0, (int)step * 2), new Point((int)X, (int)step * 2));
                    g.DrawLine(bPen, new Point(0, (int)step), new Point((int)X, (int)step));

                    g.DrawString((0.25 / zoom * 100).ToString("0.0") + "%", fontNormal, bBrush, X + 6, (step * 3) - 14);
                    g.DrawString((0.50 / zoom * 100).ToString("0.0") + "%", fontNormal, bBrush, X + 6, (step * 2) - 14);
                    g.DrawString((0.75 / zoom * 100).ToString("0.0") + "%", fontNormal, bBrush, X + 6, step - 14);

                    float sumProduct = 0, sumService = 0;
                    float sumProductAgr = 0, sumServiceAgr = 0;
                    var gp = new GraphicsPath();
                    var gpShapes = new GraphicsPath(FillMode.Winding);
                    var gpTextMonth = new GraphicsPath();
                    var gpTextWeek = new GraphicsPath();
                    var gpLines = new GraphicsPath();
                    var gpLinesMonth = new GraphicsPath();
                    var gpLinesYear = new GraphicsPath();
                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        int I = d;
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageData)dt.Rows[d][2];
                        int product = store.product;
                        int service = store.service;
                        sumProduct += product;
                        sumService += service;
                        sumProductAgr += product;
                        sumServiceAgr += service;

                        if (agr.Sjekk(date))
                        {
                            p3 = new PointF(Hstep * (I + 1), Y);
                            p4 = new PointF(Hstep * (I - agr.dager + 1), Y);

                            if (sumProductAgr > 0)
                            {
                                p1 = new PointF(Hstep * (I + 1), Y - (Vstep * sumProductAgr));
                                p2 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sumProductAgr));
                                gp.StartFigure();
                                gp.AddLine(p4, p2);
                                gp.AddLine(p2, p1);
                                gp.AddLine(p1, p3);
                                g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new RectangleF(p2, new SizeF(Hstep * agr.dager, Vstep * sumProductAgr)));
                                g.DrawString(sumProductAgr.ToString(), fontNormal, new SolidBrush(Color.Black), p2);
                            }
                            if (sumServiceAgr > 0)
                            {
                                p1 = new PointF(Hstep * (I + 1), Y - (Vstep * sumServiceAgr));
                                p2 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sumServiceAgr));
                                gp.StartFigure();
                                gp.AddLine(p4, p2);
                                gp.AddLine(p2, p1);
                                gp.AddLine(p1, p3);
                                g.FillRectangle(new SolidBrush(tjenColor), new RectangleF(p2, new SizeF(Hstep * agr.dager, Vstep * sumServiceAgr)));
                                g.DrawString(sumServiceAgr.ToString(), fontNormal, new SolidBrush(Color.Black), p2);
                            }
                            if (agr.dager > 27)
                            {
                                var percentMonth = CalcPercent(sumServiceAgr, sumProductAgr);
                                gpTextMonth.AddString(Math.Round(percentMonth * 100, 0).ToString("0") + "%", new FontFamily("Verdana"), (int)FontStyle.Regular, 32 * dpi,
                                    new PointF(Hstep * (I - 20), Y - offsetY), StringFormat.GenericDefault);
                            }

                            if (agr.dager <= 27 && agr.dager >= 7)
                            {
                                var percentMonth = CalcPercent(sumServiceAgr, sumProductAgr);
                                gpTextMonth.AddString(Math.Round(percentMonth * 100, 0).ToString("0") + "%", new FontFamily("Verdana"), (int)FontStyle.Regular, 32 * dpi,
                                    new PointF(Hstep * (I - 5), Y - offsetY), StringFormat.GenericDefault);
                            }

                            sumProductAgr = 0;
                            sumServiceAgr = 0;
                        }
                        if (main.appConfig.graphAdvanced && argAgr == "dag" && sk == "")
                        {
                            int sumacc = 0;
                            if (!DBNull.Value.Equals(dt.Rows[d][dt.Columns.Count - 1]))
                            {
                                StorageTop s = (StorageTop)dt.Rows[d]["TOP"];
                                List<StorageSelger> selgere = s.selgere;
                                for (int i = 0; i < selgere.Count; i++)
                                {
                                    if (i == 0)
                                        toppselgere.Add(selgere[i].selger); // Legg til den beste selgeren
                                    sumacc += selgere[i].antall;
                                    p6 = new PointF(Hstep * I, Y - (Vstep * sumacc));
                                    g.FillRectangle(new SolidBrush(GetSelgerkodeFarge(selgere[i].selger)), new RectangleF(p6, new SizeF(Hstep, Vstep * selgere[i].antall)));
                                    if (Vstep > 20)
                                        g.DrawString(selgere[i].selger.Substring(0, 3), fontNormal, new SolidBrush(Color.Black), p6);
                                }
                            }
                        }

                        if (date.DayOfWeek == DayOfWeek.Monday && Hstep > 20)
                        {
                            p1 = new PointF(Hstep * I, Y);
                            p2 = new PointF(Hstep * I, Y + offsetY);
                            p3 = new PointF(Hstep * I, Y + 7);
                            var weekNo = FormMain.norway.Calendar.GetWeekOfYear(date, FormMain.norway.DateTimeFormat.CalendarWeekRule, FormMain.norway.DateTimeFormat.FirstDayOfWeek);
                            gpTextWeek.AddString("Uke " + weekNo.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 18 * dpi, p3, StringFormat.GenericDefault);
                            gpLines.StartFigure();
                            gpLines.AddLine(p1, p2);
                        }
                        if (date.DayOfWeek == DayOfWeek.Sunday && Hstep > 20 && argAgr == "dag")
                        {
                            p1 = new PointF((Hstep * I) + (Hstep / 6), Y - 90);
                            p2 = new PointF((Hstep * I) + (Hstep / 6), Y - 60);
                            g.DrawString(date.ToString("MMM"), fontNormal, new SolidBrush(Color.Black), p1);
                            g.DrawString(date.ToString("dd."), fontNormal, new SolidBrush(Color.Black), p2);
                        }
                        if (date.Day == 1)
                        {
                            gpShapes.StartFigure();
                            gpShapes.AddRectangle(new RectangleF(new PointF(Hstep * I + 2, Y + 2), new SizeF(75, offsetY - 4)));
                            gpTextMonth.AddString(date.ToString("MMM"), new FontFamily("Verdana"), (int)FontStyle.Regular, 32 * dpi, new PointF(Hstep * I, Y), StringFormat.GenericDefault);
                            gpLinesMonth.StartFigure();
                            gpLinesMonth.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                        if (date.DayOfYear == 1)
                        {
                            gpLinesYear.StartFigure();
                            gpLinesYear.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                    }
                    g.DrawPath(pen1, gp);

                    List<float> listPercent = new List<float> { };
                    float prev = 0, sum = 0;

                    for (int i = fav.Count; i-- > 0; )
                    {
                        gp = new GraphicsPath();
                        sum = 0; prev = 0;
                        float sumTjen = 0, percent = 0;
                        for (int d = 0; d < dt.Rows.Count; d++)
                        {
                            int I = d;
                            DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                            var store = (StorageData)dt.Rows[d][i + 2];
                            int value = store.product;
                            int valueTjen = store.service;
                            sum += value;
                            sumTjen += valueTjen;
                            percent = CalcPercent(sumTjen, sum);
                            if ((firstdate.Month == date.Month && firstdate.Day < 8) || lastdate.Month == date.Month || lastdate.Month != date.Month && firstdate.Month != date.Month || (!main.appConfig.graphHitrateMTD && screen))
                            {
                                p2 = new PointF((Hstep * (I + 1)), Y - (Y * percent * zoom));
                                p1 = new PointF((Hstep * I), Y - (Y * prev * zoom));
                                gp.AddLine(p1, p2);
                            }
                            prev = percent;
                            if (date.AddDays(1).Day == 1 && !(firstdate.Month == 5 && firstdate.Day == 1) && !FormMain.datoPeriodeVelger && dt.Rows.Count != d + 1 && !(!main.appConfig.graphHitrateMTD && screen))
                            {
                                sumTjen = 0;
                                sum = 0;
                                percent = 0;
                                prev = 0;
                                gp.StartFigure();
                            }
                        }
                        g.DrawPath(new Pen(Color.White, 12 * dpi), gp);
                        g.DrawPath(new Pen(FormMain.favColors[i], 8 * dpi), gp);
                        g.DrawLine(new Pen(FormMain.favColors[i], 8 * dpi), X, Y - (Y * percent * zoom), argX, Y - (Y * percent * zoom));
                        PaintLabelPercent(g, percent, X, Y, offsetY, offsetY, zoom, dpi);
                        listPercent.Add(percent);
                        gp.Dispose();
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new RectangleF(0, Y, argX, offsetY));
                    g.DrawPath(pen1, gpLines);
                    g.DrawPath(new Pen(Color.Green, 3 * dpi), gpLinesMonth);
                    g.DrawPath(new Pen(Color.Red, 3 * dpi), gpLinesYear);
                    g.FillPath(new SolidBrush(Color.Black), gpTextWeek);
                    g.FillPath(new SolidBrush(Color.White), gpShapes);
                    g.FillPath(new SolidBrush(Color.Black), gpTextMonth);

                    toppselgere = toppselgere.Distinct().ToList(); // fjern duplikater
                    listPercent.Reverse(); // reverser listen

                    int height = 0;
                    if (argY > 600 && main.appConfig.graphAdvanced && sk == "")
                        height += toppselgere.Count * fontHeight;
                    if (argY > 400 && sk == "")
                        height += fav.Count * fontHeight;
                    if (!main.appConfig.graphAdvanced || argY <= 600)
                        height += fontHeight;

                    g.DrawLine(pen1, new Point(0, (int)Y), new Point(argX, (int)Y)); // understrek
                    g.DrawLine(pen1, new Point((int)X, (int)Y), new Point((int)X, 0)); // slutt strek
                    if (!screen)
                        g.DrawRectangle(pen1, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1))); // ramme
                    else
                        g.DrawRectangle(new Pen(Color.Black, 1), new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1))); // ramme

                    using (var gpSkilt = new GraphicsPath()) // skilt skygge
                    {
                        gpSkilt.AddRectangle(new Rectangle(12, 12, 347, 47 + height));
                        using (PathGradientBrush _Brush = new PathGradientBrush(gpSkilt))
                        {
                            _Brush.WrapMode = WrapMode.Clamp;
                            ColorBlend _ColorBlend = new ColorBlend(3);
                            _ColorBlend.Colors = new Color[]{Color.Transparent, 
                           Color.FromArgb(180, Color.DimGray), 
                           Color.FromArgb(180, Color.DimGray)};
                            _ColorBlend.Positions = new float[] { 0f, .1f, 1f };
                            _Brush.InterpolationColors = _ColorBlend;
                            g.FillPath(_Brush, gpSkilt);
                        }
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(8, 8, 340, 40 + height)); // skilt
                    g.DrawRectangle(pen1, new Rectangle(new Point(8, 8), new Size(340, 40 + height))); // skilt ramme
                    string s2 = argAgr;
                    if (s2 == "måned")
                        s2 = "mnd";
                    g.DrawString(argTitle + " solgt pr " + s2, fontNormal, new SolidBrush(Color.Black), 44, 15);
                    g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new Rectangle(18, 18, boxLength, boxLength));
                    g.DrawRectangle(pen1, new Rectangle(18, 18, boxLength, boxLength));

                    int add = 0;
                    if (!main.appConfig.graphAdvanced || argY <= 600)
                    {
                        add += fontHeight;
                        g.DrawString("Tjenester", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.FillRectangle(new SolidBrush(tjenColor), new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                        g.DrawRectangle(pen1, new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                    }

                    if (argY > 600 && main.appConfig.graphAdvanced && sk == "")
                    {
                        for (int i = 0; i < toppselgere.Count; i++)
                        {
                            add += fontHeight;
                            g.DrawString(toppselgere[i], fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            g.FillRectangle(new SolidBrush(GetSelgerkodeFarge(toppselgere[i])), new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                            g.DrawRectangle(pen1, new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));

                        }
                    }

                    if (argY > 400 && sk == "")
                    {
                        for (int i = 0; i < fav.Count; i++)
                        {
                            add += fontHeight;
                            g.DrawString(main.avdeling.Get(fav[i]) + ": " + Math.Round(listPercent[i] * 100, 2).ToString("0.00") + " %", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            g.DrawLine(new Pen(FormMain.favColors[i], 8 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                        }
                    }

                }
                catch
                {
                }
            }
            return b;
        }

        private Bitmap DrawToBitmapOversiktChunk(int argX, int argY, DataTable dt, string argTitle, string argCaption, string argAgr, float argMax = 15, float argMaxWeek = 40, float argMaxMonth = 100, float argPercent = 1, bool screen = false)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;
                    if (screen)
                        dpi = main.appConfig.graphScreenDPI;
                    int fontHeight = Convert.ToInt32(29 * dpi);
                    int boxLength = Convert.ToInt32(22 * dpi);
                    int fontSepHeight = Convert.ToInt32((fontHeight / 5) * dpi);

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (dt.Rows.Count == 0)
                    {
                        g.DrawString("Mangler transaksjoner!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];
                    int offsetY = Convert.ToInt32(40 * dpi);
                    int offsetX = Convert.ToInt32(150 * dpi);
                    float X = argX - offsetX;
                    float Y = argY - offsetY;
                    int pSize = dt.Rows.Count;
                    float Hstep = X / pSize;

                    if (argAgr == "")
                    {
                        if (Hstep > (40 * dpi))
                            argAgr = "dag";
                        else if (Hstep <= (40 * dpi) && Hstep >= (8 * dpi))
                            argAgr = "uke";
                        else if (Hstep < (8 * dpi))
                            argAgr = "måned";
                    }
                    Logg.Debug("Hstpe: " + Hstep + " | argAgr: " + argAgr);

                    float Vstep = Y / (argMax * 1.1f);
                    if (argAgr == "uke")
                        Vstep = Y / (argMaxWeek * 1.1f);
                    if (argAgr == "måned")
                        Vstep = Y / (argMaxMonth * 1.1f);

                    PointF p1, p2, p3, p4, p6;
                    Pen pen1 = new Pen(Color.Black, 3 * dpi);
                    Pen pen2 = new Pen(Color.Green, 6 * dpi);
                    Pen bPen = new Pen(Color.LightGray, 3 * dpi);
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Font fontBig = new Font("Helvetica", 20 * dpi, FontStyle.Bold);
                    Font fontNormal = new Font("Helvetica", 18 * dpi, FontStyle.Regular);
                    Font fontSmall = new Font("Helvetica", 13 * dpi, FontStyle.Regular);
                    string[] strTjenester = new string[] { "Finans", "TA", "Strom", "Tjen" };
                    Color[] colorTjenester = new Color[] { System.Drawing.ColorTranslator.FromHtml("#f5954e"),
                        System.Drawing.ColorTranslator.FromHtml("#6699ff"),
                        System.Drawing.ColorTranslator.FromHtml("#FAF39E"),
                        System.Drawing.ColorTranslator.FromHtml("#80c34a") };
                    int columnsCount = dt.Columns.Count;
                    DagSjekk agr = new DagSjekk(argAgr);

                    g.DrawString(main.avdeling.Get(main.appConfig.Avdeling) + " " + argCaption,
                        new Font("Verdana", 36 * dpi, FontStyle.Bold), new SolidBrush(Color.LightGray), 330, 0);

                    var step = Y / 4;
                    g.DrawLine(bPen, new Point(0, (int)step * 3), new Point((int)X, (int)step * 3));
                    g.DrawLine(bPen, new Point(0, (int)step * 2), new Point((int)X, (int)step * 2));
                    g.DrawLine(bPen, new Point(0, (int)step), new Point((int)X, (int)step));

                    g.DrawString("2,5%", fontNormal, bBrush, X + 6, (step * 3) - 14);
                    g.DrawString("5,0%", fontNormal, bBrush, X + 6, (step * 2) - 14);
                    g.DrawString("7,5%", fontNormal, bBrush, X + 6, step - 14);

                    float sumValue = 0, sumFinans = 0, sumTa = 0, sumStrom = 0, sumTjen = 0;
                    var gp = new GraphicsPath();
                    var gpLines = new GraphicsPath();
                    var gpShapes = new GraphicsPath(FillMode.Winding);
                    var gpTextMonth = new GraphicsPath();
                    var gpTextWeek = new GraphicsPath();
                    var gpLinesMonth = new GraphicsPath();
                    var gpLinesYear = new GraphicsPath();
                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        int I = d;
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageButikk)dt.Rows[d][2];
                        var tjenester = (StorageTjenester)dt.Rows[d][columnsCount - 1];
                        sumValue += store.btokr;
                        sumFinans += tjenester.finans;
                        sumTa += tjenester.ta;
                        sumStrom += tjenester.strom;
                        sumTjen += tjenester.tjen;

                        if (agr.Sjekk(date))
                        {
                            p1 = new PointF(Hstep * (I + 1), Y - (Vstep * sumValue));
                            p2 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sumValue));
                            p3 = new PointF(Hstep * (I + 1), Y);
                            p4 = new PointF(Hstep * (I - agr.dager + 1), Y);

                            gp.StartFigure();
                            gp.AddLine(p4, p2);
                            gp.AddLine(p2, p1);
                            gp.AddLine(p1, p3);

                            g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new RectangleF(p2, new SizeF(Hstep * agr.dager, Vstep * sumValue)));
                            g.DrawString(ForkortTallEnkel(sumValue), fontNormal, new SolidBrush(Color.Black), p2);

                            float sum = 0;
                            sum += sumFinans;
                            p6 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sum));
                            g.FillRectangle(new SolidBrush(colorTjenester[0]), new RectangleF(p6, new SizeF(Hstep * agr.dager, Vstep * sumFinans)));
                            sum += sumTa;
                            p6 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sum));
                            g.FillRectangle(new SolidBrush(colorTjenester[1]), new RectangleF(p6, new SizeF(Hstep * agr.dager, Vstep * sumTa)));
                            sum += sumStrom;
                            p6 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sum));
                            g.FillRectangle(new SolidBrush(colorTjenester[2]), new RectangleF(p6, new SizeF(Hstep * agr.dager, Vstep * sumStrom)));
                            sum += sumTjen;
                            p6 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sum));
                            g.FillRectangle(new SolidBrush(colorTjenester[3]), new RectangleF(p6, new SizeF(Hstep * agr.dager, Vstep * sumTjen)));

                            sumValue = 0; sumFinans = 0; sumTa = 0; sumStrom = 0; sumTjen = 0;
                        }
                        if (date.DayOfWeek == DayOfWeek.Monday && Hstep > 15)
                        {
                            p1 = new PointF(Hstep * I, Y);
                            p2 = new PointF(Hstep * I, Y + offsetY);
                            p3 = new PointF(Hstep * I, Y + 7);
                            var weekNo = FormMain.norway.Calendar.GetWeekOfYear(date, FormMain.norway.DateTimeFormat.CalendarWeekRule, FormMain.norway.DateTimeFormat.FirstDayOfWeek);
                            gpTextWeek.AddString("Uke " + weekNo.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 18 * dpi, p3, StringFormat.GenericDefault);
                            gpLines.StartFigure();
                            gpLines.AddLine(p1, p2);
                        }
                        if (date.DayOfWeek == DayOfWeek.Sunday && Hstep > 20 && argAgr == "dag")
                        {
                            p1 = new Point((int)(Hstep * I) + (int)(Hstep / 6), (int)Y - 90);
                            p2 = new Point((int)(Hstep * I) + (int)(Hstep / 6), (int)Y - 60);
                            g.DrawString(date.ToString("MMM"), fontNormal, new SolidBrush(Color.Black), p1);
                            g.DrawString(date.ToString("dd."), fontNormal, new SolidBrush(Color.Black), p2);
                        }
                        if (date.Day == 1)
                        {
                            gpShapes.StartFigure();
                            gpShapes.AddRectangle(new RectangleF(new PointF(Hstep * I + 2, Y + 2), new SizeF(75, offsetY - 4)));
                            gpTextMonth.AddString(date.ToString("MMM"), new FontFamily("Verdana"), (int)FontStyle.Regular, 32 * dpi, new PointF(Hstep * I, Y), StringFormat.GenericDefault);
                            gpLinesMonth.StartFigure();
                            gpLinesMonth.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                        if (date.DayOfYear == 1)
                        {
                            gpLinesYear.StartFigure();
                            gpLinesYear.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                    }
                    g.DrawPath(pen1, gp);

                    float percentFinans = 0, percentTa = 0, percentStrom = 0, percentTjen = 0;
                    var gpFinans = new GraphicsPath();
                    var gpTa = new GraphicsPath();
                    var gpStrom = new GraphicsPath();
                    var gpTjen = new GraphicsPath();
                    float sumInntjen = 0, sumOmset = 0, sumOmsetExMva = 0;
                    sumFinans = 0; sumTa = 0; sumStrom = 0; sumTjen = 0;
                    float percentFinansPrev = 0, percentTaPrev = 0, percentStromPrev = 0, percentTjenPrev = 0;
                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        int I = d;
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageTjenester)dt.Rows[d][columnsCount - 1];
                        var butikk = (StorageButikk)dt.Rows[d][2];
                        sumInntjen += butikk.btokr;
                        sumOmset += butikk.salgspris;
                        sumOmsetExMva += butikk.salgexmva;
                        sumFinans += store.finans;
                        percentFinans = CalcPercent(sumFinans, sumInntjen);
                        sumTa += store.taOmset;
                        percentTa = CalcPercent(sumTa, sumOmsetExMva);
                        sumStrom += store.strom;
                        percentStrom = CalcPercent(sumStrom, sumInntjen);
                        sumTjen += store.tjen;
                        percentTjen = CalcPercent(sumTjen, sumInntjen);
                        if ((firstdate.Month == date.Month && firstdate.Day < 8) || lastdate.Month == date.Month || lastdate.Month != date.Month && firstdate.Month != date.Month || (!main.appConfig.graphHitrateMTD && screen))
                        {
                            p2 = new PointF((Hstep * (I + 1)), Y - (Y * percentFinans * 10));
                            p1 = new PointF((Hstep * I), Y - (Y * percentFinansPrev * 10));
                            gpFinans.AddLine(p1, p2);
                            p2 = new PointF((Hstep * (I + 1)), Y - (Y * percentTa * 10));
                            p1 = new PointF((Hstep * I), Y - (Y * percentTaPrev * 10));
                            gpTa.AddLine(p1, p2);
                            p2 = new PointF((Hstep * (I + 1)), Y - (Y * percentStrom * 10));
                            p1 = new PointF((Hstep * I), Y - (Y * percentStromPrev * 10));
                            gpStrom.AddLine(p1, p2);
                            p2 = new PointF((Hstep * (I + 1)), Y - (Y * percentTjen * 10));
                            p1 = new PointF((Hstep * I), Y - (Y * percentTjenPrev * 10));
                            gpTjen.AddLine(p1, p2);
                        }
                        percentFinansPrev = percentFinans;
                        percentTaPrev = percentTa;
                        percentStromPrev = percentStrom;
                        percentTjenPrev = percentTjen;
                        if (date.AddDays(1).Day == 1 && !(firstdate.Month == 5 && firstdate.Day == 1) && !FormMain.datoPeriodeVelger && dt.Rows.Count != d + 1 && !(!main.appConfig.graphHitrateMTD && screen))
                        {
                            sumFinans = 0; sumTa = 0; sumStrom = 0; sumTjen = 0; sumInntjen = 0; sumOmset = 0; sumOmsetExMva = 0;
                            percentFinans = 0; percentTa = 0; percentStrom = 0; percentTjen = 0;
                            percentFinansPrev = 0; percentTaPrev = 0; percentStromPrev = 0; percentTjenPrev = 0;
                            gpFinans.StartFigure(); gpTa.StartFigure(); gpStrom.StartFigure(); gpTjen.StartFigure();
                        }
                    }
                    // Finans path
                    g.DrawPath(new Pen(Color.White, 12 * dpi), gpFinans);
                    g.DrawPath(new Pen(colorTjenester[0], 8 * dpi), gpFinans);
                    g.DrawLine(new Pen(colorTjenester[0], 8 * dpi), X, Y - (Y * percentFinans * 10), argX, Y - (Y * percentFinans * 10));
                    PaintLabelPercent(g, percentFinans, X, Y, offsetX, offsetY, 10, dpi);
                    // TA path
                    g.DrawPath(new Pen(Color.White, 12 * dpi), gpTa);
                    g.DrawPath(new Pen(colorTjenester[1], 8 * dpi), gpTa);
                    g.DrawLine(new Pen(colorTjenester[1], 8 * dpi), X, Y - (Y * percentTa * 10), argX, Y - (Y * percentTa * 10));
                    PaintLabelPercent(g, percentTa, X, Y, offsetX, offsetY, 10, dpi);
                    // Strøm path
                    g.DrawPath(new Pen(Color.White, 12 * dpi), gpStrom);
                    g.DrawPath(new Pen(colorTjenester[2], 8 * dpi), gpStrom);
                    g.DrawLine(new Pen(colorTjenester[2], 8 * dpi), X, Y - (Y * percentStrom * 10), argX, Y - (Y * percentStrom * 10));
                    PaintLabelPercent(g, percentStrom, X, Y, offsetX, offsetY, 10, dpi);
                    // RTG/SA path
                    g.DrawPath(new Pen(Color.White, 12 * dpi), gpTjen);
                    g.DrawPath(new Pen(colorTjenester[3], 8 * dpi), gpTjen);
                    g.DrawLine(new Pen(colorTjenester[3], 8 * dpi), X, Y - (Y * percentTjen * 10), argX, Y - (Y * percentTjen * 10));
                    PaintLabelPercent(g, percentTjen, X, Y, offsetX, offsetY, 10, dpi);

                    g.FillRectangle(new SolidBrush(Color.White), new RectangleF(0, Y, argX, offsetY));
                    g.DrawPath(pen1, gpLines);
                    g.DrawPath(new Pen(Color.Green, 3 * dpi), gpLinesMonth);
                    g.DrawPath(new Pen(Color.Red, 3 * dpi), gpLinesYear);
                    g.FillPath(new SolidBrush(Color.Black), gpTextWeek);
                    g.FillPath(new SolidBrush(Color.White), gpShapes);
                    g.FillPath(new SolidBrush(Color.Black), gpTextMonth);

                    g.DrawLine(pen1, new Point(0, (int)Y), new Point(argX, (int)Y));
                    g.DrawLine(pen1, new Point((int)X, (int)Y), new Point((int)X, 0));
                    g.DrawRectangle(pen1, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1)));

                    int height = 0;
                    if (argY > 200)
                        height += fontHeight * 4;

                    using (var gpSkilt = new GraphicsPath()) // skilt skygge
                    {
                        gpSkilt.AddRectangle(new Rectangle(12, 12, 307, 47 + height));
                        using (PathGradientBrush _Brush = new PathGradientBrush(gpSkilt))
                        {
                            _Brush.WrapMode = WrapMode.Clamp;
                            ColorBlend _ColorBlend = new ColorBlend(3);
                            _ColorBlend.Colors = new Color[]{Color.Transparent, 
                           Color.FromArgb(180, Color.DimGray), 
                           Color.FromArgb(180, Color.DimGray)};
                            _ColorBlend.Positions = new float[] { 0f, .1f, 1f };
                            _Brush.InterpolationColors = _ColorBlend;
                            g.FillPath(_Brush, gpSkilt);
                        }
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(8, 8, 300, 40 + height));
                    g.DrawRectangle(pen1, new Rectangle(new Point(8, 8), new Size(300, 40 + height)));
                    string s2 = argAgr;
                    if (s2 == "måned")
                        s2 = "mnd";
                    g.DrawString("Fortjeneste pr " + s2, fontNormal, new SolidBrush(Color.Black), 44, 15);
                    g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new Rectangle(18, 18, boxLength, boxLength));
                    g.DrawRectangle(pen1, new Rectangle(18, 18, boxLength, boxLength));

                    int add = 0;
                    if (argY > 200)
                    {
                        add += fontHeight;
                        g.DrawString("Finans:", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.DrawString(Math.Round(percentFinans * 100, 2).ToString("0.00") + " %", fontNormal, new SolidBrush(Color.Black), 130, 15 + add);
                        g.DrawLine(new Pen(colorTjenester[0], 10 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                        add += fontHeight;
                        g.DrawString("TA:", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.DrawString(Math.Round(percentTa * 100, 2).ToString("0.00") + " %", fontNormal, new SolidBrush(Color.Black), 130, 15 + add);
                        g.DrawLine(new Pen(colorTjenester[1], 10 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                        add += fontHeight;
                        g.DrawString("Strøm:", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.DrawString(Math.Round(percentStrom * 100, 2).ToString("0.00") + " %", fontNormal, new SolidBrush(Color.Black), 130, 15 + add);
                        g.DrawLine(new Pen(colorTjenester[2], 10 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                        add += fontHeight;
                        g.DrawString("RTG/SA:", fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.DrawString(Math.Round(percentTjen * 100, 2).ToString("0.00") + " %", fontNormal, new SolidBrush(Color.Black), 130, 15 + add);
                        g.DrawLine(new Pen(colorTjenester[3], 10 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                    }
                }
                catch
                {
                }
            }
            return b;
        }

        private Bitmap DrawToBitmapButikkChunk(int argX, int argY, DataTable dt, string argTitle, string argCaption, string argAgr, float argMax = 15, float argMaxWeek = 40, float argMaxMonth = 100, float argPercent = 1, bool screen = false, float argSumTotal = 1)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;
                    if (screen)
                        dpi = main.appConfig.graphScreenDPI;
                    int fontHeight = Convert.ToInt32(29 * dpi);
                    int boxLength = Convert.ToInt32(22 * dpi);
                    int fontSepHeight = Convert.ToInt32((fontHeight / 5) * dpi);

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (dt.Rows.Count == 0)
                    {
                        g.DrawString("Mangler transaksjoner!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    DateTime firstdate = DateTime.Now;
                    DateTime lastdate = DateTime.Now;
                    firstdate = (DateTime)dt.Rows[0][1];
                    lastdate = (DateTime)dt.Rows[dt.Rows.Count - 1][1];
                    int offsetY = Convert.ToInt32(40 * dpi);
                    int offsetX = Convert.ToInt32(150 * dpi);
                    float X = argX - offsetX;
                    float Y = argY - offsetY;
                    int pSize = dt.Rows.Count;
                    float Hstep = X / pSize;

                    if (argAgr == "")
                    {
                        if (Hstep > (40 * dpi))
                            argAgr = "dag";
                        else if (Hstep <= (40 * dpi) && Hstep >= (8 * dpi))
                            argAgr = "uke";
                        else if (Hstep < (8 * dpi))
                            argAgr = "måned";
                    }

                    float Vstep = Y / (argMax * 1.1f);
                    if (argAgr == "uke")
                        Vstep = Y / (argMaxWeek * 1.1f);
                    if (argAgr == "måned")
                        Vstep = Y / (argMaxMonth * 1.1f);

                    float VstepA = Y / (argSumTotal + (argSumTotal / 10));

                    PointF p1, p2, p3, p4, p6;
                    Pen bPen = new Pen(Color.LightGray, 3 * dpi);
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen1 = new Pen(Color.Black, 3 * dpi);
                    Pen pen2 = new Pen(Color.Green, 6 * dpi);
                    Font fontBig = new Font("Helvetica", 20 * dpi, FontStyle.Bold);
                    Font fontNormal = new Font("Helvetica", 18 * dpi, FontStyle.Regular);
                    Font fontSmall = new Font("Helvetica", 13 * dpi, FontStyle.Regular);
                    List<string> toppselgere = new List<string> { };
                    DagSjekk agr = new DagSjekk(argAgr);

                    g.DrawString(main.avdeling.Get(main.appConfig.Avdeling) + " " + lastdate.ToString("MMMM yyyy"),
                        new Font("Verdana", 36 * dpi, FontStyle.Bold), new SolidBrush(Color.LightGray), 400, 0);

                    var step = Y / 4;
                    g.DrawLine(bPen, new Point(0, (int)step * 3), new Point((int)X, (int)step * 3));
                    g.DrawLine(bPen, new Point(0, (int)step * 2), new Point((int)X, (int)step * 2));
                    g.DrawLine(bPen, new Point(0, (int)step), new Point((int)X, (int)step));

                    var verticalSum = argMax / 4;
                    if (argAgr == "uke")
                        verticalSum = argMaxWeek / 4;
                    if (argAgr == "måned")
                        verticalSum = argMaxMonth / 4;
                    g.DrawString(ForkortTall(verticalSum + (verticalSum / 10)), fontNormal, bBrush, X + 6, (step * 3) - 14);
                    g.DrawString(ForkortTall(verticalSum * 2 + (verticalSum / 10)), fontNormal, bBrush, X + 6, (step * 2) - 14);
                    g.DrawString(ForkortTall(verticalSum * 3 + (verticalSum / 10)), fontNormal, bBrush, X + 6, step - 14);

                    float sumValue = 0;
                    var gp = new GraphicsPath();
                    var gpShapes = new GraphicsPath(FillMode.Winding);
                    var gpTextMonth = new GraphicsPath();
                    var gpTextWeek = new GraphicsPath();
                    var gpLines = new GraphicsPath();
                    var gpLinesMonth = new GraphicsPath();
                    var gpLinesYear = new GraphicsPath();
                    for (int d = 0; d < dt.Rows.Count; d++)
                    {
                        int I = d;
                        DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                        var store = (StorageButikk)dt.Rows[d][2];
                        sumValue += store.btokr;

                        if (agr.Sjekk(date))
                        {
                            p1 = new PointF(Hstep * (I + 1), Y - (Vstep * sumValue));
                            p2 = new PointF(Hstep * (I - agr.dager + 1), Y - (Vstep * sumValue));
                            p3 = new PointF(Hstep * (I + 1), Y);
                            p4 = new PointF(Hstep * (I - agr.dager + 1), Y);

                            g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new RectangleF(p2, new SizeF(Hstep * agr.dager, Vstep * sumValue)));
                            g.DrawString(ForkortTallEnkel(sumValue), fontNormal, new SolidBrush(Color.Black), p2);

                            if (main.appConfig.graphAdvanced && argAgr == "dag")
                            {
                                int sumacc = 0;
                                if (!DBNull.Value.Equals(dt.Rows[d]["TOP"]))
                                {
                                    StorageTop s = (StorageTop)dt.Rows[d]["TOP"];
                                    List<StorageSelger> selgere = s.selgere;
                                    for (int i = 0; i < selgere.Count; i++)
                                    {
                                        if (i == 0)
                                            toppselgere.Add(selgere[i].selger); // Legg til den beste selgeren
                                        sumacc += selgere[i].antall;
                                        if ((selgere[i].antall * Vstep) > 20)
                                        {
                                            p6 = new PointF(Hstep * I, Y - (Vstep * sumacc));
                                            g.FillRectangle(new SolidBrush(GetSelgerkodeFarge(selgere[i].selger)), new RectangleF(p6, new SizeF(Hstep, Vstep * selgere[i].antall)));
                                            g.DrawString(selgere[i].selger.Substring(0, 3), fontNormal, new SolidBrush(Color.Black), p6);
                                        }
                                    }
                                }
                            }

                            gp.StartFigure();
                            gp.AddLine(p4, p2);
                            gp.AddLine(p2, p1);
                            gp.AddLine(p1, p3);
                            sumValue = 0;
                        }

                        if (date.DayOfWeek == DayOfWeek.Monday && Hstep > 20)
                        {
                            p1 = new PointF(Hstep * I, Y);
                            p2 = new PointF(Hstep * I, Y + offsetY);
                            p3 = new PointF(Hstep * I, Y + 7);
                            var weekNo = FormMain.norway.Calendar.GetWeekOfYear(date, FormMain.norway.DateTimeFormat.CalendarWeekRule, FormMain.norway.DateTimeFormat.FirstDayOfWeek);
                            gpTextWeek.AddString("Uke " + weekNo.ToString(), new FontFamily("Verdana"), (int)FontStyle.Regular, 18 * dpi, p3, StringFormat.GenericDefault);
                            gpLines.StartFigure();
                            gpLines.AddLine(p1, p2);
                        }
                        if (date.DayOfWeek == DayOfWeek.Sunday && Hstep > 20 && argAgr == "dag")
                        {
                            p1 = new Point((int)(Hstep * I) + (int)(Hstep / 6), (int)Y - 90);
                            p2 = new Point((int)(Hstep * I) + (int)(Hstep / 6), (int)Y - 60);
                            g.DrawString(date.ToString("MMM"), fontNormal, new SolidBrush(Color.Black), p1);
                            g.DrawString(date.ToString("dd."), fontNormal, new SolidBrush(Color.Black), p2);
                        }
                        if (date.Day == 1)
                        {
                            gpShapes.StartFigure();
                            gpShapes.AddRectangle(new RectangleF(new PointF(Hstep * I + 2, Y + 2), new SizeF(75, offsetY - 4)));
                            gpTextMonth.AddString(date.ToString("MMM"), new FontFamily("Verdana"), (int)FontStyle.Regular, 32 * dpi, new PointF(Hstep * I, Y), StringFormat.GenericDefault);
                            gpLinesMonth.StartFigure();
                            gpLinesMonth.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                        if (date.DayOfYear == 1)
                        {
                            gpLinesYear.StartFigure();
                            gpLinesYear.AddLine(new PointF(Hstep * I, 0), new PointF(Hstep * I, argY));
                        }
                    }
                    g.DrawPath(pen1, gp);

                    List<int> favSum = new List<int> { };
                    float sumCompare = 0, prevCompare = 0;
                    for (int i = Favoritter.Count; i-- > 0; )
                    {
                        var gpSum = new GraphicsPath();
                        var gpSumCompare = new GraphicsPath();
                        float sum = 0, prev = 0;
                        for (int d = 0; d < dt.Rows.Count; d++)
                        {
                            int I = d;
                            DateTime date = Convert.ToDateTime(dt.Rows[d][1]);
                            var store = (StorageButikk)dt.Rows[d][i + 2];
                            sum += store.btokr;

                            if ((firstdate.Month == date.Month && firstdate.Day < 8) || lastdate.Month == date.Month || lastdate.Month != date.Month && firstdate.Month != date.Month || (!main.appConfig.graphHitrateMTD && screen))
                            {
                                p2 = new PointF((Hstep * (I + 1)), Y - (VstepA * sum));
                                p1 = new PointF((Hstep * I), Y - (VstepA * prev));
                                gpSum.AddLine(p1, p2);

                                if (main.appConfig.rankingCompareLastyear > 0 && i == 0)
                                {
                                    sumCompare += store.ifjor_btokr;
                                    p2 = new PointF((Hstep * (I + 1)), Y - (VstepA * sumCompare));
                                    p1 = new PointF((Hstep * I), Y - (VstepA * prevCompare));
                                    gpSumCompare.AddLine(p1, p2);
                                    prevCompare = sumCompare;
                                }
                            }

                            prev = sum;
                            if (date.AddDays(1).Day == 1 && !(firstdate.Month == 5 && firstdate.Day == 1) && !FormMain.datoPeriodeVelger && dt.Rows.Count != d + 1 && !(!main.appConfig.graphHitrateMTD && screen))
                            {
                                sum = 0; sumCompare = 0; prev = 0; prevCompare = 0;
                                gpSum.StartFigure(); gpSumCompare.StartFigure();
                            }
                        }

                        if (main.appConfig.rankingCompareLastyear > 0 && i == 0)
                        {
                            g.DrawPath(new Pen(Color.White, 12 * dpi), gpSumCompare);
                            g.DrawPath(new Pen(Color.Gray, 8 * dpi), gpSumCompare);
                            g.DrawLine(new Pen(Color.Gray, 8 * dpi), X, Y - (VstepA * sumCompare), argX, Y - (VstepA * sumCompare));
                            PaintLabelTall(g, sumCompare, X, Y, offsetX, offsetY, VstepA, dpi);
                        }
                        g.DrawPath(new Pen(Color.White, 12 * dpi), gpSum);
                        g.DrawPath(new Pen(FormMain.favColors[i], 8 * dpi), gpSum);
                        g.DrawLine(new Pen(FormMain.favColors[i], 8 * dpi), X, Y - (VstepA * sum), argX, Y - (VstepA * sum));
                        PaintLabelTall(g, sum, X, Y, offsetX, offsetY, VstepA, dpi);
                        favSum.Add(Convert.ToInt32(sum));
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new RectangleF(0, Y, argX, offsetY));
                    g.DrawPath(pen1, gpLines);
                    g.DrawPath(new Pen(Color.Green, 3 * dpi), gpLinesMonth);
                    g.DrawPath(new Pen(Color.Red, 3 * dpi), gpLinesYear);
                    g.FillPath(new SolidBrush(Color.Black), gpTextWeek);
                    g.FillPath(new SolidBrush(Color.White), gpShapes);
                    g.FillPath(new SolidBrush(Color.Black), gpTextMonth);

                    g.DrawLine(pen1, new Point(0, (int)Y), new Point(argX, (int)Y));
                    g.DrawLine(pen1, new Point((int)X, (int)Y), new Point((int)X, 0)); // slutt strek
                    g.DrawRectangle(pen1, new Rectangle(new Point(0, 0), new Size(argX - 1, argY - 1)));

                    toppselgere = toppselgere.Distinct().ToList(); // fjern duplikater
                    favSum.Reverse();

                    int height = 0;
                    if (argY > 400)
                        height += Favoritter.Count * fontHeight;
                    if (argY > 600 && main.appConfig.graphAdvanced)
                        height += toppselgere.Count * fontHeight;
                    if (main.appConfig.rankingCompareLastyear > 0 && sumCompare > 0)
                        height += fontHeight;

                    using (var gpSkilt = new GraphicsPath()) // skilt skygge
                    {
                        gpSkilt.AddRectangle(new Rectangle(12, 12, 347, 47 + height));
                        using (PathGradientBrush _Brush = new PathGradientBrush(gpSkilt))
                        {
                            _Brush.WrapMode = WrapMode.Clamp;
                            ColorBlend _ColorBlend = new ColorBlend(3);
                            _ColorBlend.Colors = new Color[]{Color.Transparent, 
                           Color.FromArgb(180, Color.DimGray), 
                           Color.FromArgb(180, Color.DimGray)};
                            _ColorBlend.Positions = new float[] { 0f, .1f, 1f };
                            _Brush.InterpolationColors = _ColorBlend;
                            g.FillPath(_Brush, gpSkilt);
                        }
                    }
                    g.FillRectangle(new SolidBrush(Color.White), new Rectangle(8, 8, 340, 40 + height));
                    g.DrawRectangle(pen1, new Rectangle(new Point(8, 8), new Size(340, 40 + height)));
                    string s2 = argAgr;
                    if (s2 == "måned")
                        s2 = "mnd";
                    g.DrawString("Fortjeneste pr " + s2, fontNormal, new SolidBrush(Color.Black), 44, 15);
                    g.FillRectangle(new SolidBrush(Color.AntiqueWhite), new Rectangle(18, 15, boxLength, boxLength));
                    g.DrawRectangle(pen1, new Rectangle(18, 15, boxLength, boxLength));

                    int add = 0;
                    if (argY > 600 && main.appConfig.graphAdvanced)
                    {
                        for (int i = 0; i < toppselgere.Count; i++)
                        {
                            add += fontHeight;
                            g.DrawString(toppselgere[i], fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            g.FillRectangle(new SolidBrush(GetSelgerkodeFarge(toppselgere[i])), new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                            g.DrawRectangle(pen1, new Rectangle(18, 15 + fontSepHeight + add, boxLength, boxLength));
                        }
                    }

                    if (argY > 400)
                    {
                        for (int i = 0; i < Favoritter.Count; i++)
                        {
                            add += fontHeight;
                            var prosent = Math.Round(CalcPercent(favSum[i], favSum[0]) * 100, 2).ToString("0.00") + " %";
                            g.DrawString(main.avdeling.Get(Favoritter[i]) + ": " + prosent, fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                            g.DrawLine(new Pen(FormMain.favColors[i], 8 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                        }
                    }
                    if (main.appConfig.rankingCompareLastyear > 0 && sumCompare > 0)
                    {
                        add += fontHeight;
                        var prosent = Math.Round(CalcPercent(sumCompare, favSum[0]) * 100, 2).ToString("0.00") + " %";
                        g.DrawString("Fortjeneste i fjor: " + prosent, fontNormal, new SolidBrush(Color.Black), 44, 15 + add);
                        g.DrawLine(new Pen(Color.Gray, 8 * dpi), new Point(18, 15 + (fontSepHeight * 3) + add), new Point(40, 15 + (fontSepHeight * 3) + add));
                    }
                }
                catch
                {
                }
            }
            return b;
        }

    }
}
