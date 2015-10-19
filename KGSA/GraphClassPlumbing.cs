using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Serialization;


namespace KGSA
{
    public partial class GraphClass : IDisposable
    {
        private void PaintLabelTall(Graphics g, float argTall, float X, float Y, float offsetX, float offsetY, float Vstep, float dpi = 1)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddString(ForkortTall(argTall), new FontFamily("Helvetica"),
            (int)FontStyle.Bold, 28f * dpi, new PointF(X + (offsetX / 4), Y - (Vstep * argTall) - 30), StringFormat.GenericDefault);

            for (int i = 1; i < 8; ++i)
            {
                Pen pen = new Pen(Color.FromArgb(32, 255, 255, 255), i);
                pen.LineJoin = LineJoin.Round;
                g.DrawPath(pen, path);
                pen.Dispose();
            }

            SolidBrush brush = new SolidBrush(Color.FromArgb(0, 0, 0));
            g.FillPath(brush, path);
        }

        private void PaintLabelPercent(Graphics g, float argPercent, float X, float Y, float offsetX, float offsetY, float zoom = 10, float dpi = 1)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddString(Math.Round(argPercent * 100, 2).ToString("0.00") + " %", new FontFamily("Helvetica"),
            (int)FontStyle.Bold, 28f * dpi, new PointF(X + (offsetX / 4), Y - (Y * argPercent * zoom) - 30), StringFormat.GenericDefault);

            for (int i = 1; i < 8; ++i)
            {
                Pen pen = new Pen(Color.FromArgb(32, 255, 255, 255), i);
                pen.LineJoin = LineJoin.Round;
                g.DrawPath(pen, path);
                pen.Dispose();
            }

            SolidBrush brush = new SolidBrush(Color.FromArgb(0, 0, 0));
            g.FillPath(brush, path);
        }

        private string GetVarekodeSqlString(string argKat = "")
        {
            try
            {
                List<VarekodeList> list;
                if (!String.IsNullOrEmpty(argKat) && argKat != "KnowHow")
                    list = main.appConfig.varekoder.Where(item => item.kategori == argKat && item.inclhitrate == true).ToList();
                else
                    list = main.appConfig.varekoder.Where(item => item.synlig == true && item.inclhitrate == true).ToList();

                IEnumerable<string> listvk = list.Select(x => x.kode).Distinct().ToList();

                string sql = "";

                foreach (var varekode in listvk)
                    if (sql.Length > 0)
                        sql += " OR Varekode = '" + varekode + "'";
                    else
                        sql += " Varekode = '" + varekode + "'";

                return sql;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
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
            catch
            {
                return 0;
            }
        }

        MD5 md5 = MD5.Create();
        private Color GetSelgerkodeFarge(string skArg)
        {
            var value = skArg;
            if (skArg.StartsWith("MAR") || skArg.StartsWith("VER")) // Legg til kollisjoner her
                value = skArg.Substring(1, skArg.Length - 1);
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(value));
            var color = Color.FromArgb(hash[0], hash[1], hash[2]);
            Color changed = Tint(color, Color.AntiqueWhite, 0.5M);

            return changed;
        }

        private static Color Tint(Color source, Color tint, decimal alpha)
        {
            //(tint -source)*alpha + source
            int red = Convert.ToInt32(((tint.R - source.R) * alpha + source.R));
            int blue = Convert.ToInt32(((tint.B - source.B) * alpha + source.B));
            int green = Convert.ToInt32(((tint.G - source.G) * alpha + source.G));
            return Color.FromArgb(255, red, green, blue);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //dispose managed ressources
                    dataData.Dispose();
                    dataNettbrett.Dispose();
                    dataAudioVideo.Dispose();
                    dataTele.Dispose();
                    dataButikk.Dispose();
                    dataOversikt.Dispose();
                }
            }
            //dispose unmanaged ressources
            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private string ForkortTall(double arg)
        {
            var r = Math.Round(arg / 1000d, 1);
            if (r < 2)
                return arg.ToString("0.0");

            var b = Math.Round(arg / 1000000d, 1);
            if (b < 2)
                return r.ToString("0.0") + "k";

            return b.ToString("0.0") + "m";

        }

        private string ForkortTallEnkel(double arg)
        {
            var r = Math.Round(arg / 1000d, 1);
            if (r < 2)
                return arg.ToString("0");

            var b = Math.Round(arg / 1000000d, 1);
            if (b < 2)
                return r.ToString("0") + "k";

            return b.ToString("0.0") + "m";

        }
    }

    public class StorageData
    {
        public int product { get; set; }
        public int service { get; set; }
        public StorageData(int argProduct, int argService)
        {
            this.product = argProduct;
            this.service = argService;
        }
    }

    public class StorageButikk
    {
        public int btokr { get; set; }
        public int salgspris { get; set; }
        public int salgexmva { get; set; }
        public int ifjor_btokr { get; set; }
        public StorageButikk(int argBtokr, int argSalgspris, int argSalgexmva)
        {
            this.btokr = argBtokr;
            this.salgspris = argSalgspris;
            this.salgexmva = argSalgexmva;
        }
    }

    public class StorageTjenester
    {
        public int finans { get; set; }
        public int ta { get; set; }
        public int strom { get; set; }
        public int tjen { get; set; }
        public int taOmset { get; set; }
        public StorageTjenester(int argFinans, int argTa, int argStrom, int argTjen, int argTaOmset)
        {
            this.finans = argFinans;
            this.ta = argTa;
            this.strom = argStrom;
            this.tjen = argTjen;
            this.taOmset = argTaOmset;
        }
    }

    public class StorageTop
    {
        public List<StorageSelger> selgere { get; set; }
        public StorageTop()
        {
            selgere = new List<StorageSelger> { };
        }
    }

    public class StorageSelger
    {
        public string selger { get; set; }
        public int antall { get; set; }

        public StorageSelger(string selger, int antall)
        {
            this.selger = selger;
            this.antall = antall;
        }
    }

    public class DagSjekk
    {
        private string agr;
        private DateTime dato;
        public int dager { get; set; }

        public DagSjekk(string argument)
        {
            this.agr = argument;
        }

        public bool Sjekk(DateTime date)
        {
            this.dato = date;
            this.dager = Dager();
            if (agr == "dag")
                return true;
            if (agr == "uke" && dato.DayOfWeek == DayOfWeek.Sunday)
                return true;
            if (agr == "måned" && dato.AddDays(1).Day == 1)
                return true;
            return false;
        }

        private int Dager()
        {
            if (agr == "dag")
                return 1;
            if (agr == "uke")
                return 7;
            if (agr == "måned")
                return FormMain.GetLastDayOfMonth(dato.AddDays(-1)).Day;
            return 1;
        }
    }
}
