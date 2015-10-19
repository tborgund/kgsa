using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TopGraph
    {
        FormMain main;
        private DataTable dataButikk { get; set; }
        private DataTable dataData { get; set; }
        private DataTable dataAudioVideo { get; set; }
        private DataTable dataTele { get; set; }
        public DateTime dato = FormMain.rangeMin;
        public DateTime datoFra = FormMain.rangeMin;
        private int maxButikk = 0;
        private int maxData = 0;
        private int maxAudioVideo = 0;
        private int maxTele = 0;
        private int dager = 0;
        private string currentGraph = "";

        public TopGraph(FormMain form)
        {
            this.main = form;
            datoFra = main.appConfig.dbTo.AddMonths(-1);
            dato = main.appConfig.dbTo;
            dager = (dato - datoFra).Days;
            dataButikk = GetHistTable();
            dataData = GetHistTable();
            dataAudioVideo = GetHistTable();
            dataTele = GetHistTable();
        }

        public DataTable UpdateGraph(string katArg)
        {
            try
            {
                if (katArg == "Data")
                {
                    if (main.appConfig.dbTo.Date != dato.Date || dataData.Rows.Count == 0)
                        dataData = UpdateDataTable(katArg);
                    currentGraph = katArg;
                    return dataData;
                }
                else if (katArg == "AudioVideo")
                {
                    if (main.appConfig.dbTo.Date != dato.Date || dataAudioVideo.Rows.Count == 0)
                        dataAudioVideo = UpdateDataTable(katArg);
                    currentGraph = katArg;
                    return dataAudioVideo;
                }
                else if (katArg == "Tele")
                {
                    if (main.appConfig.dbTo.Date != dato.Date || dataTele.Rows.Count == 0)
                        dataTele = UpdateDataTable(katArg);
                    currentGraph = katArg;
                    return dataTele;
                }
                else
                {
                    if (main.appConfig.dbTo.Date != dato.Date || dataButikk.Rows.Count == 0)
                        dataButikk = UpdateDataTable(katArg);
                    currentGraph = "Butikk";
                    return dataButikk;
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        private void PaintLabel(Graphics g, string text, float X, float Y)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddString(text, new FontFamily("Helvetica"), (int)FontStyle.Bold, 28f, new PointF(1.5f, 1.5f), StringFormat.GenericDefault);
            for (int i = 1; i < 8; ++i)
            {
                Pen pen = new Pen(Color.FromArgb(32, 230, 230, 230), i);
                pen.LineJoin = LineJoin.Round;
                g.DrawPath(pen, path);
                pen.Dispose();
            }

            SolidBrush brush = new SolidBrush(Color.FromArgb(0, 0, 0));
            g.FillPath(brush, path);
            var gpLabel = new GraphicsPath();
            gpLabel.AddString(main.avdeling.Get(main.appConfig.Avdeling) + " " + currentGraph, new FontFamily("Helvetica"), (int)FontStyle.Bold, 28f, new PointF(0, 0), StringFormat.GenericDefault);
            g.FillPath(new SolidBrush(Color.White), gpLabel);
            g.DrawPath(new Pen(Color.Black, 0.7f), gpLabel);
        }

        public Bitmap DrawToBitmapTop(Bitmap b, int argX, int argY)
        {
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    DataTable dt = GetHistTable();
                    if (currentGraph == "Butikk")
                        dt = dataButikk;
                    if (currentGraph == "Data")
                        dt = dataData;
                    if (currentGraph == "AudioVideo")
                        dt = dataAudioVideo;
                    if (currentGraph == "Tele")
                        dt = dataTele;

                    float X = argX;
                    float Y = argY;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);

                    if (dt.Rows.Count == 0)
                        return b;

                    int pSize = dt.Rows.Count;
                    if (pSize == 0)
                        pSize = 1;
                    float Hstep = X / pSize;
                    if (Hstep > 80)
                        Hstep = 80;

                    var max = 8;
                    if (currentGraph == "Butikk")
                        max = maxButikk;
                    else if (currentGraph == "Data")
                        max = maxData;
                    else if (currentGraph == "AudioVideo")
                        max = maxAudioVideo;
                    else if (currentGraph == "Tele")
                        max = maxTele;

                    float Vstep = Y / max;

                    PointF p1, p2, p3, p4;
                    var pen = new Pen(Color.Black, 1);
                    Color productColor = Color.AntiqueWhite;
                    Color serviceColor = ColorTranslator.FromHtml("#80c34a");
                    var font = new Font("Helvetica", 10, FontStyle.Regular);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        int I = (int)dt.Rows[i][0];
                        DateTime d = Convert.ToDateTime(dt.Rows[i][1]);
                        int D = (int)dt.Rows[i][2];
                        int T = (int)dt.Rows[i][3];

                        var bData = new SolidBrush(productColor);
                        var bTjen = new SolidBrush(serviceColor);

                        if (d == FormMain.highlightDate)
                        {
                            bTjen.Color = Tint(bTjen.Color, Color.Yellow, 0.3M);
                            bData.Color = Tint(bData.Color, Color.Yellow, 0.3M);
                        }

                        if (dato.Month != d.Month)
                        {
                            bData.Color = Tint(bData.Color, Color.LightGray, 0.9M);
                            bTjen.Color = Tint(bTjen.Color, Color.LightGray, 0.9M);
                        }
                        p1 = new PointF(X - (Hstep * I), Y - (Vstep * D));
                        p2 = new PointF(X - (Hstep * (I + 1)), Y - (Vstep * D));
                        p3 = new PointF(X - (Hstep * I), Y);
                        p4 = new PointF(X - (Hstep * (I + 1)), Y);

                        g.FillRectangle(bData, new RectangleF(p2, new SizeF(Hstep, Vstep * D)));
                        g.DrawLine(pen, new Point(Convert.ToInt32(X - (Hstep * I)), Convert.ToInt32(Y - (Vstep * D))),
                            new Point(Convert.ToInt32(X - (Hstep * (I + 1))), Convert.ToInt32(Y - (Vstep * D)))); // topp
                        g.DrawLine(pen, p3, p1); // høyre
                        g.DrawLine(pen, p4, p2); // venstre

                        p1 = new PointF(X - (Hstep * I), Y - (Vstep * T));
                        p2 = new PointF(X - (Hstep * (I + 1)), Y - (Vstep * T));
                        p3 = new PointF(X - (Hstep * I), Y);
                        p4 = new PointF(X - (Hstep * (I + 1)), Y);
                        g.FillRectangle(bTjen, new RectangleF(p2, new SizeF(Hstep, Vstep * T)));
                        g.DrawLine(pen, new Point(Convert.ToInt32(X - (Hstep * I)), Convert.ToInt32(Y - (Vstep * T))),
                            new Point(Convert.ToInt32(X - (Hstep * (I + 1))), Convert.ToInt32(Y - (Vstep * T)))); // topp
                        g.DrawLine(pen, p3, p1); // høyre
                        g.DrawLine(pen, p4, p2); // venstre
                        g.DrawString(ForkortTallEnkel(D), font, new SolidBrush(Color.Black),
                            new PointF(X - (Hstep * (I + 1)), Y - (Vstep * D)));
                        g.DrawString(ForkortTallEnkel(T), font, new SolidBrush(Color.Black),
                            new PointF(X - (Hstep * I), Y - (Vstep * T) + 1),
                            new StringFormat(StringFormatFlags.DirectionRightToLeft));

                        if (d.DayOfWeek == DayOfWeek.Sunday && D == 0 && Hstep > 6)
                        {
                            g.DrawString(d.ToString("MMM"), font, new SolidBrush(Color.Black),
                                new PointF(X - (Hstep * (I + 1) - (Hstep / 4)), Y - 20));
                            g.DrawString(d.ToString("dd."), font, new SolidBrush(Color.Black),
                                new PointF(X - (Hstep * (I + 1) - (Hstep / 4)), Y - 35));
                            g.DrawLine(pen, new PointF(X - (Hstep * (I + 1) - (Hstep / 4) + 2), Y),
                                new PointF(X - (Hstep * (I + 1) - (Hstep / 4) + 2), Y - 38));
                        }

                        if (d.Day == 1 && Hstep <= 6)
                        {
                            g.DrawString(d.ToString("MMM"), font, new SolidBrush(Color.Black),
                                new PointF(X - (Hstep * (I + 1) - (Hstep / 4)), Y - 20));
                            g.DrawLine(pen, new PointF(X - (Hstep * (I + 1) - (Hstep / 4) + 2), Y),
                                new PointF(X - (Hstep * (I + 1) - (Hstep / 4) + 2), Y - 38));
                        }

                        if (d.Day == 1)
                            g.DrawLine(pen, new PointF(X - (Hstep * (I + 1)), 0), new PointF(X - (Hstep * (I + 1)), Y + 50));
                    }

                    PaintLabel(g, main.avdeling.Get(main.appConfig.Avdeling) + " " + currentGraph, X, Y);
                }
                catch
                {
                }
            }
            return b;
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

        private static Color Tint(Color source, Color tint, decimal alpha)
        {
            //(tint -source)*alpha + source
            int red = Convert.ToInt32(((tint.R - source.R) * alpha + source.R));
            int blue = Convert.ToInt32(((tint.B - source.B) * alpha + source.B));
            int green = Convert.ToInt32(((tint.G - source.G) * alpha + source.G));
            return Color.FromArgb(255, red, green, blue);
        }

        private string GetVarekodeSqlString(string argKat = "")
        {
            try
            {
                List<VarekodeList> list;
                if (!String.IsNullOrEmpty(argKat))
                    list = main.appConfig.varekoder.Where(item => item.kategori == argKat).ToList();
                else
                    list = main.appConfig.varekoder.ToList();

                IEnumerable<string> listvk = list.Where(item => item.synlig == true).Select(x => x.kode).Distinct().ToList();
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

        private static DataTable GetHistTable()
        {
            var table = new DataTable();
            table.Columns.Add("Index", typeof(int));
            table.Columns.Add("Dato", typeof(DateTime));
            table.Columns.Add("Produkt", typeof(int));
            table.Columns.Add("Tjeneste", typeof(int));
            return table;
        }

        private DataTable UpdateDataTable(string katArg = "Data")
        {
            try
            {
                DataTable dt = GetHistTable();

                string command = "";
                string product = "";

                if (katArg == "Data")
                    product = "Varegruppe = '531' OR Varegruppe = '533'";
                else if (katArg == "AudioVideo")
                    product = "Varegruppe = '224'";
                else if (katArg == "Tele")
                    product = "Varegruppe = '431'";

                if (katArg == "Butikk")
                    command = "SELECT SUM(Salgspris) AS Salgspris, SUM(Btokr) AS Btokr, Dato FROM tblSalg " +
                        "WHERE (Avdeling = '" + main.appConfig.Avdeling + "') AND (Dato >= '" + datoFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dato.ToString("yyy-MM-dd") + "') GROUP BY Dato";
                else
                    command = "SELECT SUM(CASE WHEN " + product + " THEN Antall ELSE 0 END) AS Product, " +
                        "SUM(CASE WHEN " + GetVarekodeSqlString(katArg) + " THEN Antall ELSE 0 END) AS Service, " +
                        "Dato FROM tblSalg WHERE (Avdeling = '" + main.appConfig.Avdeling + "') AND (Dato >= '" + datoFra.ToString("yyy-MM-dd") + "') AND (Dato <= '" + dato.ToString("yyy-MM-dd") + "') GROUP BY Dato";

                DataTable sqlce = main.database.GetSqlDataTable(command);

                int max = 0;
                int count = sqlce.Rows.Count;

                if (count > 0)
                {
                    for (int d = 0; d < dager; d++)
                    {
                        DataRow dtRow = dt.NewRow();
                        dtRow[0] = d;
                        dtRow[1] = dato.AddDays(-d);
                        dtRow[2] = 0;
                        dtRow[3] = 0;
                        for (int i = 0; i < sqlce.Rows.Count; i++)
                        {
                            if (Convert.ToDateTime(sqlce.Rows[i][2]) == dato.AddDays(-d))
                            {
                                int sum = Convert.ToInt32(sqlce.Rows[i][0]);
                                if (max < sum)
                                    max = sum;
                                dtRow[2] = sum;
                                dtRow[3] = Convert.ToInt32(sqlce.Rows[i][1]);
                            }

                        }
                        dt.Rows.Add(dtRow);
                    }

                    if (katArg == "Butikk")
                        maxButikk = max;
                    if (katArg == "Data")
                        maxData = max;
                    if (katArg == "AudioVideo")
                        maxAudioVideo = max;
                    if (katArg == "Tele")
                        maxTele = max;
                    return dt;
                }
                return dt;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

    }
}
