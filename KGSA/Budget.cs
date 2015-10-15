using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class BudgetObj
    {
        FormMain main;

        public BudgetObj(FormMain form)
        {
            this.main = form;
        }

        public void UpdateBudgetSelgerkoder(int budgetId)
        {
            try
            {
                var currentBudgetInfo = GetBudgetInfo(budgetId);

                if (currentBudgetInfo != null)
                {
                    var sk = main.sKoder.GetBudgetSelgerkoder(currentBudgetInfo.kategori);

                    var command = new SqlCeCommand("SELECT * FROM tblBudgetSelger WHERE BudgetId = " + budgetId, main.connection);
                    var da = new SqlCeDataAdapter(command);
                    var ds = new DataSet();
                    da.Fill(ds, "tblBudgetSelger");
                    var ca = new SqlCeCommandBuilder(da);
                    foreach (string selger in sk)
                    {
                        var foundAuthors = ds.Tables["tblBudgetSelger"].Select("Selgerkode = '" + selger + "'");
                        if (foundAuthors.Length == 0)
                        {
                            DataRow dRow = ds.Tables["tblBudgetSelger"].NewRow();
                            dRow["BudgetId"] = budgetId;
                            dRow["Selgerkode"] = selger;
                            dRow["Timer"] = 0;
                            dRow["Dager"] = 0;
                            dRow["Multiplikator"] = 1;
                            dRow["Comment"] = "";
                            ds.Tables["tblBudgetSelger"].Rows.Add(dRow);
                        }
                    }
                    da.Update(ds, "tblBudgetSelger");

                    command = new SqlCeCommand("SELECT * FROM tblBudgetTimer WHERE BudgetId = " + budgetId, main.connection);
                    da = new SqlCeDataAdapter(command);
                    ds = new DataSet();
                    da.Fill(ds, "tblBudgetTimer");
                    ca = new SqlCeCommandBuilder(da);
                    foreach (string selger in sk)
                    {
                        var foundAuthors = ds.Tables["tblBudgetTimer"].Select("Selgerkode = '" + selger + "'");
                        if (foundAuthors.Length == 0)
                        {
                            DataRow dRow = ds.Tables["tblBudgetTimer"].NewRow();
                            dRow["BudgetId"] = budgetId;
                            dRow["Selgerkode"] = selger;
                            ds.Tables["tblBudgetTimer"].Rows.Add(dRow);
                        }
                    }
                    da.Update(ds, "tblBudgetTimer");
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void DeleteBudget(int budgetId)
        {
            try
            {
                using (var con = new SqlCeConnection(FormMain.SqlConStr))
                {
                    con.Open();
                    SqlCeCommand cmd = con.CreateCommand();
                    cmd.CommandText = "DELETE FROM tblBudget WHERE Id = " + budgetId;
                    cmd.ExecuteScalar();
                    cmd.CommandText = "DELETE FROM tblBudgetSelger WHERE BudgetId = " + budgetId;
                    cmd.ExecuteScalar();
                    cmd.CommandText = "DELETE FROM tblBudgetTimer WHERE BudgetId = " + budgetId;
                    cmd.ExecuteScalar();
                    con.Close();
                }
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public void SumWorkHoursAndDays(BudgetInfo budget_info)
        {
            try
            {
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                var command = new SqlCeCommand("SELECT * FROM tblBudgetSelger WHERE BudgetId = " + budget_info.budget_id, con);
                var da = new SqlCeDataAdapter(command);
                var ds = new DataSet();
                da.Fill(ds, "tblBudgetTimer");
                var ca = new SqlCeCommandBuilder(da);

                DataTable dtTimer = main.database.GetSqlDataTable("SELECT * FROM tblBudgetTimer WHERE BudgetId = " + budget_info.budget_id);

                if (dtTimer.Rows.Count == 0)
                    return;

                for (int i = 0; i < dtTimer.Rows.Count; i++)
                {
                    for (int b = 0; b < ds.Tables["tblBudgetTimer"].Rows.Count; b++)
                    {
                        if (ds.Tables["tblBudgetTimer"].Rows[b]["Selgerkode"].ToString() == dtTimer.Rows[i]["Selgerkode"].ToString())
                        {
                            decimal timer = 0, dager = 0;

                            for (int c = 3; c < dtTimer.Columns.Count; c++)
                            {
                                if (!DBNull.Value.Equals(dtTimer.Rows[i][c]))
                                {
                                    decimal t = Convert.ToDecimal(dtTimer.Rows[i][c]);
                                    timer += t;
                                    dager++;
                                }
                            }

                            ds.Tables["tblBudgetTimer"].Rows[b]["Timer"] = timer;
                            ds.Tables["tblBudgetTimer"].Rows[b]["Dager"] = dager;

                        }
                    }
                }

                da.Update(ds, "tblBudgetTimer");

                // Oppdater timeantall oppdaterings tid
                con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();
                command = new SqlCeCommand("SELECT * FROM tblBudget WHERE Id = " + budget_info.budget_id, con);
                da = new SqlCeDataAdapter(command);
                ds = new DataSet();
                da.Fill(ds, "tblBudget");
                ca = new SqlCeCommandBuilder(da);
                if (ds.Tables["tblBudget"].Rows.Count > 0)
                    ds.Tables["tblBudget"].Rows[0]["Updated"] = DateTime.Now;
                else
                    Logg.Log("Fikk ikke satt tidspunkt for timeantall oppdatering.", Color.Red);
                da.Update(ds, "tblBudget");

                con.Close();
                con.Dispose();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        public string GetProductColor(BudgetType type)
        {
            if (type == BudgetType.Vinnprodukt)
                return "#4cad88";
            if (type == BudgetType.Finans)
                return "#f5954e";
            if (type == BudgetType.TA)
                return "#6699ff";
            if (type == BudgetType.Strom)
                return "#FAF39E";
            if (type == BudgetType.Rtgsa)
                return "#80c34a";
            if (type == BudgetType.Acc)
                return "#f36565";
            if (type == BudgetType.Inntjening)
                return "#ff5b2e";
            if (type == BudgetType.Omsetning)
                return "#ff5b2e";

             return "#f5954e";
        }

        public DataTable GetAllBudgets()
        {
            DataTable dt = main.database.GetSqlDataTable("SELECT * FROM tblBudget WHERE Avdeling = " + main.appConfig.Avdeling);
            return dt;
        }

        public BudgetInfo GetBudgetInfo(DateTime date, BudgetCategory cat)
        {
            return GetBudgetInternal(date, 0, cat);
        }

        public BudgetInfo GetBudgetInfo(int budgetId)
        {
            return GetBudgetInternal(DateTime.Now, budgetId);
        }

        private BudgetInfo GetBudgetInternal(DateTime date, int budgetId = 0, BudgetCategory cat = BudgetCategory.None)
        {
            try
            {
                string sql = "";
                
                if (budgetId == 0)
                    sql = "SELECT * FROM tblBudget WHERE Avdeling = " + main.appConfig.Avdeling + " AND Kategori = '" + BudgetCategoryClass.TypeToName(cat) + "' AND datepart(Month, Date) = " + date.Month + " AND datepart(Year, Date) = " + date.Year;
                else
                    sql = "SELECT * FROM tblBudget WHERE Id = " + budgetId;

                DataTable dt = main.database.GetSqlDataTable(sql);

                if (dt.Rows.Count > 0)
                {
                    var _budgetObj = new BudgetInfo();

                    // Legg till generell info
                    budgetId = Convert.ToInt32(dt.Rows[0]["Id"].ToString());
                    _budgetObj.budget_id = budgetId;
                    _budgetObj.date = Convert.ToDateTime(dt.Rows[0]["Date"].ToString());
                    _budgetObj.dager = Convert.ToInt32(dt.Rows[0]["Dager"].ToString());
                    _budgetObj.avdeling = Convert.ToInt32(dt.Rows[0]["Avdeling"].ToString());
                    _budgetObj.kategori = BudgetCategoryClass.NameToType(dt.Rows[0]["Kategori"].ToString());
                    _budgetObj.omsetning = Convert.ToDecimal(dt.Rows[0]["Omsetning"].ToString());
                    _budgetObj.inntjening = Convert.ToDecimal(dt.Rows[0]["Inntjening"].ToString());
                    _budgetObj.margin = Convert.ToDecimal(dt.Rows[0]["Margin"].ToString());

                    _budgetObj.ta = Convert.ToDecimal(dt.Rows[0]["TA"].ToString());
                    _budgetObj.ta_type = ConvertToBudgetType(dt.Rows[0]["TA_Type"].ToString());
                    _budgetObj.strom = Convert.ToDecimal(dt.Rows[0]["Strom"].ToString());
                    _budgetObj.strom_type = ConvertToBudgetType(dt.Rows[0]["Strom_Type"].ToString());
                    _budgetObj.finans = Convert.ToDecimal(dt.Rows[0]["Finans"].ToString());
                    _budgetObj.finans_type = ConvertToBudgetType(dt.Rows[0]["Finans_Type"].ToString());
                    _budgetObj.rtgsa = Convert.ToDecimal(dt.Rows[0]["Rtgsa"].ToString());
                    _budgetObj.rtgsa_type = ConvertToBudgetType(dt.Rows[0]["Rtgsa_Type"].ToString());
                    _budgetObj.acc = Convert.ToDecimal(dt.Rows[0]["Acc"].ToString());
                    _budgetObj.acc_type = ConvertToBudgetType(dt.Rows[0]["Acc_Type"].ToString());
                    _budgetObj.vinn = Convert.ToDecimal(dt.Rows[0]["Vinn"].ToString());
                    _budgetObj.vinn_type = ConvertToBudgetType(dt.Rows[0]["Vinn_Type"].ToString());

                    _budgetObj.updated = Convert.ToDateTime(dt.Rows[0]["Updated"].ToString());

                    // Legg til selgere under dette budsjettet
                    DataTable dtSel = main.database.GetSqlDataTable("SELECT Selgerkode, Timer, Dager, Multiplikator FROM tblBudgetSelger WHERE BudgetId = " + budgetId);
                    if (dtSel.Rows.Count > 0)
                    {
                        List<BudgetSelger> list = new List<BudgetSelger>() { };
                        int timer = 0, dager = 0;
                        for(int i = 0; i < dtSel.Rows.Count; i++)
                        {
                            BudgetSelger bselger = new BudgetSelger();
                            bselger.selgerkode = dtSel.Rows[i]["Selgerkode"].ToString();

                            bselger.timer = Convert.ToInt32(dtSel.Rows[i]["Timer"]);
                            timer += bselger.timer;
                            bselger.dager = Convert.ToInt32(dtSel.Rows[i]["Dager"]);
                            dager += bselger.dager;
                            bselger.multiplikator = Convert.ToDecimal(dtSel.Rows[i]["Multiplikator"]);
                            list.Add(bselger);
                        }
                        _budgetObj.totalt_timer = timer;
                        _budgetObj.totalt_dager = dager;

                        for (int i = 0; i < list.Count; i++ )
                        {
                            decimal weight = 0;
                            if (timer > 0)
                                weight = (list[i].timer * list[i].multiplikator) / timer;

                            list[i].weight = weight;
                        }
                        _budgetObj.selgere = list;
                    }

                    // Regne ut coefficient tall for beregning av budsjett MTD
                    decimal daysElapsed = GetNumberOfOpenDaysInMonth(FormMain.GetFirstDayOfMonth(date), date);
                    decimal daysInMonth = _budgetObj.dager;
                    _budgetObj.timeElapsedCoefficient = daysElapsed / daysInMonth;
                    _budgetObj.daysElapsed = (int)daysElapsed;

                    return _budgetObj;
                }

            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
            return null;
        }

        public int GetNumberOfOpenDaysInMonth(DateTime from, DateTime to)
        {
            int days = 0;
            foreach (DateTime day in EachDay(from, to))
            {
                if (day.DayOfWeek != DayOfWeek.Sunday && !(
                    day.Date == new DateTime(from.Year, 1, 1) ||
                    day.Date == new DateTime(from.Year, 5, 1) ||
                    day.Date == new DateTime(from.Year, 5, 17) ||
                    day.Date == new DateTime(from.Year, 12, 25) ||
                    day.Date == new DateTime(from.Year, 12, 26)))
                    days++;
            }
            return days;
        }

        public IEnumerable<DateTime> EachDay(DateTime from, DateTime thru)
        {
            for (var day = from.Date; day.Date <= thru.Date; day = day.AddDays(1))
                yield return day;
        }

        private BudgetValueType ConvertToBudgetType(string value)
        {

            if (value != null)
            {
                if (value == "Poeng")
                    return BudgetValueType.Poeng;
                if (value == "Omset")
                    return BudgetValueType.Omsetning;
                if (value == "Inntjen")
                    return BudgetValueType.Inntjening;
                if (value == "SoM")
                    return BudgetValueType.SoM;
                if (value == "SoB")
                    return BudgetValueType.SoB;
                if (value == "Hitrate")
                    return BudgetValueType.Hitrate;
                if (value == "Antall")
                    return BudgetValueType.Antall;
            }

            return BudgetValueType.Inntjening;
        }

        public string GetValueTypeSuffix(BudgetValueType type)
        {
            if (type == BudgetValueType.Poeng)
                return "poeng";
            if (type == BudgetValueType.Antall)
                return "stk";
            if (type == BudgetValueType.AntallPerDag)
                return "stk /dag";
            if (type == BudgetValueType.Hitrate)
                return "%";
            if (type == BudgetValueType.Inntjening)
                return "kr";
            if (type == BudgetValueType.InntjeningPerDag)
                return "kr /dag";
            if (type == BudgetValueType.Omsetning)
                return "kr";
            if (type == BudgetValueType.OmsetningPerDag)
                return "kr /dag";
            if (type == BudgetValueType.SoB)
                return "%";
            if (type == BudgetValueType.SoM)
                return "%";
            
            return "";
        }

        public string ProductToString(BudgetType type)
        {
            if (type == BudgetType.Vinnprodukt)
                return "Vinnprodukt";
            if (type == BudgetType.Acc)
                return "Tilbehør";
            if (type == BudgetType.Finans)
                return "Finansiering";
            if (type == BudgetType.Rtgsa)
                return "RTG/SA";
            if (type == BudgetType.Strom)
                return "Norges Energi";
            if (type == BudgetType.TA)
                return "Trygghetsavtale";
            if (type == BudgetType.Inntjening)
                return "Inntjening";
            if (type == BudgetType.Omsetning)
                return "Omsetning";
            if (type == BudgetType.Kvalitet)
                return "Kvalitet";
            if (type == BudgetType.Effektivitet)
                return "Effektivitet";

            return "Ukjent";
        }

        private bool CheckDuplicate(int avdeling, BudgetCategory cat, DateTime date)
        {
            var con = new SqlCeConnection(FormMain.SqlConStr);
            con.Open();
            SqlCeCommand cmd = con.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM tblBudget WHERE Avdeling = " + avdeling + " AND Kategori = '" + BudgetCategoryClass.TypeToName(cat) + "' AND datepart(Month, Date) = " + date.Month + " AND datepart(Year, Date) = " + date.Year;
            int result = ((int)cmd.ExecuteScalar());
            con.Close();
            con.Dispose();
            if (result == 0)
                return false;
            return true;
        }

        public int AddBudget(int avdeling, BudgetCategory cat, DateTime date, int dager, decimal omsetning, decimal inntjening,
            decimal margin, decimal ta_value = 0, string ta_type = "SOB", decimal strom_value = 0, string strom_type = "Antall",
            decimal finans_value = 0, string finans_type = "SOM", decimal rtgsa_value = 0, string rtgsa_type = "Hitrate",
            decimal acc_value = 0, string acc_type = "SOM", decimal vinn_value = 0, string vinn_type = "Poeng", bool update = false, int id = 0)
        {

            try
            {
                var con = new SqlCeConnection(FormMain.SqlConStr);
                con.Open();

                if (CheckDuplicate(avdeling, cat, date) && !update)
                {
                    Logg.Log("Budsjett: Dette budsjettet finnes allerede i databasen. Slett den gamle først hvis du mener å lage en ny.");
                    return -1;
                }

                if (update)
                {
                    using (SqlCeCommand cmd = new SqlCeCommand("UPDATE tblBudget SET "
                        + "Avdeling=@Avdeling,Kategori=@Kategori,Date=@Date,Dager=@Dager,Omsetning=@Omsetning,Inntjening=@Inntjening,Margin=@Margin," +
                        "TA=@TA,TA_Type=@TA_Type,Strom=@Strom,Strom_Type=@Strom_Type,Finans=@Finans,Finans_Type=@Finans_Type,Rtgsa=@Rtgsa,Acc=@Acc,Acc_Type=@Acc_Type,Vinn=@Vinn,Vinn_Type=@Vinn_Type,Updated=@Updated" +
                         " WHERE Id=@Id", con))
                    {
                        cmd.Parameters.AddWithValue("@Id", id);
                        cmd.Parameters.AddWithValue("@Avdeling", avdeling);
                        cmd.Parameters.AddWithValue("@Kategori", BudgetCategoryClass.TypeToName(cat));
                        cmd.Parameters.AddWithValue("@Date", date);
                        cmd.Parameters.AddWithValue("@Dager", dager);
                        cmd.Parameters.AddWithValue("@Omsetning", omsetning);
                        cmd.Parameters.AddWithValue("@Inntjening", inntjening);
                        cmd.Parameters.AddWithValue("@Margin", margin);
                        cmd.Parameters.AddWithValue("@TA", ta_value);
                        cmd.Parameters.AddWithValue("@TA_Type", ta_type);
                        cmd.Parameters.AddWithValue("@Strom", strom_value);
                        cmd.Parameters.AddWithValue("@Strom_Type", strom_type);
                        cmd.Parameters.AddWithValue("@Finans", finans_value);
                        cmd.Parameters.AddWithValue("@Finans_Type", finans_type);
                        cmd.Parameters.AddWithValue("@Rtgsa", rtgsa_value);
                        cmd.Parameters.AddWithValue("@Rtgsa_Type", rtgsa_type);
                        cmd.Parameters.AddWithValue("@Acc", acc_value);
                        cmd.Parameters.AddWithValue("@Acc_Type", acc_type);
                        cmd.Parameters.AddWithValue("@Vinn", vinn_value);
                        cmd.Parameters.AddWithValue("@Vinn_Type", vinn_type);
                        cmd.Parameters.AddWithValue("@Updated", date);

                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    using (SqlCeCommand cmd = new SqlCeCommand("INSERT INTO tblBudget(Avdeling, Kategori, Date, Dager, Omsetning, Inntjening, Margin, TA, TA_Type, Strom, Strom_Type, Finans, Finans_Type, Rtgsa, Rtgsa_Type, Acc, Acc_Type, Vinn, Vinn_Type, Updated) values (@Val1, @Val2, @Val3, @Val4, @Val5, @Val6, @Val7, @Val8, @Val9, @Val10, @Val11, @Val12, @Val13, @Val14, @Val15, @Val16, @Val17, @Val18, @Val19, @Val20)", con))
                    {
                        cmd.Parameters.AddWithValue("@Val1", avdeling);
                        cmd.Parameters.AddWithValue("@Val2", BudgetCategoryClass.TypeToName(cat));
                        cmd.Parameters.AddWithValue("@Val3", date);
                        cmd.Parameters.AddWithValue("@Val4", dager);
                        cmd.Parameters.AddWithValue("@Val5", omsetning);
                        cmd.Parameters.AddWithValue("@Val6", inntjening);
                        cmd.Parameters.AddWithValue("@Val7", margin);
                        cmd.Parameters.AddWithValue("@Val8", ta_value);
                        cmd.Parameters.AddWithValue("@Val9", ta_type);
                        cmd.Parameters.AddWithValue("@Val10", strom_value);
                        cmd.Parameters.AddWithValue("@Val11", strom_type);
                        cmd.Parameters.AddWithValue("@Val12", finans_value);
                        cmd.Parameters.AddWithValue("@Val13", finans_type);
                        cmd.Parameters.AddWithValue("@Val14", rtgsa_value);
                        cmd.Parameters.AddWithValue("@Val15", rtgsa_type);
                        cmd.Parameters.AddWithValue("@Val16", acc_value);
                        cmd.Parameters.AddWithValue("@Val17", acc_type);
                        cmd.Parameters.AddWithValue("@Val18", vinn_value);
                        cmd.Parameters.AddWithValue("@Val19", vinn_type);
                        cmd.Parameters.AddWithValue("@Val20", FormMain.GetFirstDayOfMonth(date));
                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "SELECT @@IDENTITY";
                        id = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
                con.Close();
                con.Dispose();

                return id;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }

            return -1;
        }


        public void SaveBarChartImage(string argFilename, BudgetInfo budgetinfo)
        {
            Bitmap graphBitmap = DrawBarChartData(budgetinfo);
            if (graphBitmap != null)
            {
                graphBitmap.Save(argFilename, ImageFormat.Png);
                graphBitmap.Dispose();
            }
        }

        public void SaveChartImage(int argX, int argY, string argFilename, BudgetValueType type, BudgetType product, List<BudgetChartData> list)
        {
            Bitmap graphBitmap = DrawChartData(argX, argY, type, product, list);
            if (graphBitmap != null)
            {
                graphBitmap.Save(argFilename, ImageFormat.Png);
                graphBitmap.Dispose();
            }
        }

        private Bitmap DrawBarChartData(BudgetInfo budgetinfo, int argX = 820, int argY = 350)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);

                    if (budgetinfo.barchart == null)
                    {
                        g.DrawString("Mangler info!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    float X = argX;
                    float Y = argY;

                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen = new Pen(Color.Black, 3 * dpi);
                    Font fontNormal = new Font("Helvetica", 22 * dpi, FontStyle.Regular);
                    Font fontMedium = new Font("Helvetica", 20 * dpi, FontStyle.Regular);
                    Font fontSmall = new Font("Helvetica", 16 * dpi, FontStyle.Regular);
                    Color resultColor = System.Drawing.ColorTranslator.FromHtml("#ff5b2e");
                    Color resultColorDark = Tint(System.Drawing.ColorTranslator.FromHtml("#ff5b2e"), Color.Black, 0.5M);
                    Color targetColor = System.Drawing.ColorTranslator.FromHtml("#1163b0");
                    Color targetColorDark = System.Drawing.ColorTranslator.FromHtml("#004586");

                    float offsetTop = 20;
                    float barTop = 100;
                    float barBottom = 300;
                    float barHeight = 200;
                    float barWidth = 75;

                    Color resColorIntjen = Color.Green, resColorOmset = Color.Green;
                    if (budgetinfo.barchart.result_omset < 100)
                        resColorOmset = Color.Red;
                    if (budgetinfo.barchart.result_inntjen < 100)
                        resColorIntjen = Color.Red;

                    float omset = (float)budgetinfo.barchart.result_omset / 100;
                    float inntjen = (float)budgetinfo.barchart.result_inntjen / 100;

                    float omsetBarHeight = omset * 200;
                    float omsetBarTop = barTop - (omsetBarHeight - 200);

                    float inntjenBarHeight = inntjen * 200;
                    float inntjenBarTop = barTop - (inntjenBarHeight - 200);


                    g.DrawLine(new Pen(Color.Gray, 5 * dpi), new Point(75 - 50, (int)barBottom), new PointF(300 - 50, (int)barBottom));

                    g.FillRectangle(new SolidBrush(resultColor), new RectangleF(new PointF(100 - 50, omsetBarTop), new SizeF(barWidth, omsetBarHeight)));
                    g.DrawRectangle(new Pen(resultColorDark, 3 * dpi), 100 - 50, omsetBarTop, barWidth, omsetBarHeight);

                    g.FillRectangle(new SolidBrush(targetColor), new RectangleF(new PointF(200 - 50, barTop), new SizeF(barWidth, barHeight)));
                    g.DrawRectangle(new Pen(targetColorDark, 3 * dpi), 200 - 50, barTop, barWidth, barHeight);

                    g.DrawString("Omsetning:", fontNormal, new SolidBrush(Color.Black), new PointF(70 - 50, offsetTop));
                    g.DrawString(budgetinfo.barchart.result_omset + " %", fontNormal, new SolidBrush(resColorOmset), new PointF(230 - 50, offsetTop));


                    g.DrawLine(new Pen(Color.Gray, 5 * dpi), new Point(430 - 100, (int)barBottom), new PointF(655 - 100, (int)barBottom));

                    g.FillRectangle(new SolidBrush(resultColor), new RectangleF(new PointF(455 - 100, inntjenBarTop), new SizeF(barWidth, inntjenBarHeight)));
                    g.DrawRectangle(new Pen(resultColorDark, 3 * dpi), 455 - 100, inntjenBarTop, barWidth, inntjenBarHeight);

                    g.FillRectangle(new SolidBrush(targetColor), new RectangleF(new PointF(555 - 100, barTop), new SizeF(barWidth, barHeight)));
                    g.DrawRectangle(new Pen(targetColorDark, 3 * dpi), 555 - 100, barTop, barWidth, barHeight);

                    g.DrawString("Dag: " + budgetinfo.daysElapsed + " / " + budgetinfo.dager, fontMedium, new SolidBrush(Color.Black), new PointF(610, offsetTop + 50));
                    g.DrawString("Andel: " + Math.Round(budgetinfo.timeElapsedCoefficient * 100, 0) + " %", fontMedium, new SolidBrush(Color.Black), new PointF(610, offsetTop + 90));

                    g.DrawString("Inntjening:", fontNormal, new SolidBrush(Color.Black), new PointF(430 - 100, offsetTop));
                    g.DrawString(budgetinfo.barchart.result_inntjen + " %", fontNormal, new SolidBrush(resColorIntjen), new PointF(575 - 100, offsetTop));

                    g.DrawString("Resultat", fontSmall, new SolidBrush(Color.Black), new PointF(45, 300));
                    g.DrawString("Mål", fontSmall, new SolidBrush(Color.Black), new PointF(170, 300));


                    g.DrawString("Resultat", fontSmall, new SolidBrush(Color.Black), new PointF(350, 300));
                    g.DrawString("Mål", fontSmall, new SolidBrush(Color.Black), new PointF(475, 300));
                    
                }
                catch
                {
                }
            }
            return b;
        }

        private Bitmap DrawChartData(int argX, int argY, BudgetValueType type, BudgetType product, List<BudgetChartData> list)
        {
            Bitmap b = new Bitmap(argX, argY);
            using (Graphics g = Graphics.FromImage(b))
            {
                try
                {
                    float dpi = 1;

                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.White);
                    if (list.Count == 0)
                    {
                        g.DrawString("Mangler info!", new Font("Verdana", 30, FontStyle.Bold), new SolidBrush(Color.Red), 400, 0);
                        return b;
                    }

                    float offsetTop = 20;
                    float offsetLeft = 25;
                    float offsetBottom = 35;
                    float X = argX;
                    float Y = argY;

                    PointF p1, p2;
                    SolidBrush bBrush = new SolidBrush(Color.Gray);
                    Pen pen = new Pen(Color.Black, 3 * dpi);
                    Font fontNormal = new Font("Helvetica", 22 * dpi, FontStyle.Regular);
                    Color resultColor = System.Drawing.ColorTranslator.FromHtml(GetProductColor(product));
                    Color resultColorDark = Tint(System.Drawing.ColorTranslator.FromHtml(GetProductColor(product)), Color.Black, 0.5M);

                    Color targetColor = System.Drawing.ColorTranslator.FromHtml("#1163b0");
                    Color targetColorDark = System.Drawing.ColorTranslator.FromHtml("#004586");

                    int count = list.Count;
                    if (count < main.appConfig.budgetChartMinPosts)
                        count = main.appConfig.budgetChartMinPosts;

                    int maxI = 0, mI = 0;
                    for (int i = 0; i < list.Count; i++)
                    {
                        mI = Convert.ToInt32(list[i].actual);
                        if (mI > maxI)
                            maxI = mI;

                        mI = Convert.ToInt32(list[i].target);
                        if (mI > maxI)
                            maxI = mI;
                    }

                    float innject = 200 * dpi;
                    float maxW = X - (X / 8) - innject;
                    float Vstep = (Y - 100) / count;

                    int roundTo = 1000;

                    if (maxI > 1000000)
                        roundTo = 125000;
                    else if (maxI > 500000)
                        roundTo = 100000;
                    else if (maxI > 100000)
                        roundTo = 40000;
                    else if (maxI > 50000)
                        roundTo = 10000;
                    else if (maxI > 25000)
                        roundTo = 5000;
                    else if (maxI > 5000)
                        roundTo = 1000;
                    else if (maxI > 1000)
                        roundTo = 500;
                    else if (maxI > 250)
                        roundTo = 50;
                    else if (maxI > 150)
                        roundTo = 25;
                    else if (maxI > 100)
                        roundTo = 20;
                    else if (maxI > 50)
                        roundTo = 10;
                    else if (maxI > 25)
                        roundTo = 5;
                    else if (maxI > 10)
                        roundTo = 2;
                    else if (maxI > 1)
                        roundTo = 1;

                    for (int i = 0; i < 100; i++)
                    {
                        float someNumber = maxI - (roundTo * i);
                        int closest = (int)((someNumber + (0.5 * roundTo)) / roundTo) * roundTo;
                        float wI = maxW * ((float)closest / (float)maxI);
                        if (float.IsNaN(wI) || float.IsInfinity(wI))
                            wI = 1;
                        if (closest < 0)
                            break;

                        g.DrawLine(new Pen(Color.LightGray, 5 * dpi), new PointF(wI + offsetLeft + innject, 0 + offsetTop), new PointF(wI + offsetLeft + innject, Y - offsetBottom));
                        g.DrawString(ForkortTallBudget(closest) + " " + GetValueTypeSuffix(type), new Font("Verdana", 20 * dpi, FontStyle.Regular), new SolidBrush(Color.Black), wI + offsetLeft + innject, Y - offsetBottom);
                    }

                    System.Drawing.StringFormat drawFormat = new System.Drawing.StringFormat(StringFormatFlags.DirectionVertical);
                    g.DrawString(ProductToString(product), new Font("Verdana", 38 * dpi, FontStyle.Bold), new SolidBrush(Color.Gray), X - (100 * dpi), 30, drawFormat);

                    for (int i = 0; i < count; i++)
                    {
                        if (list.Count > i)
                        {
                            string selger = list[i].selgerkode;
                            decimal actual = list[i].actual;
                            decimal target = list[i].target;

                            float x = offsetLeft + innject;
                            float y = Vstep * i + offsetTop;
                            float y2 = Vstep * i + offsetTop + (Vstep / 5);

                            p1 = new PointF(x, y);
                            p2 = new PointF(x, y2);

                            float wI = maxW * ((float)actual / (float)maxI);
                            float wIT = maxW * ((float)target / (float)maxI);

                            if (float.IsNaN(wI) || float.IsInfinity(wI))
                                wI = 1;
                            if (float.IsNaN(wIT) || float.IsInfinity(wIT))
                                wIT = 1;

                            g.DrawString(main.sKoder.GetNavn(selger), fontNormal, new SolidBrush(Color.Black), new PointF(x + 10 - innject, Vstep * i + offsetTop + (Vstep / 10)));

                            if (type != BudgetValueType.Hitrate && type != BudgetValueType.SoB && type != BudgetValueType.SoM)
                            {
                                g.FillRectangle(new SolidBrush(targetColor), new RectangleF(p2, new SizeF(wIT, Vstep / 1.5f)));
                                g.DrawRectangle(new Pen(targetColorDark, 3 * dpi), x, y2, wIT, Vstep / 1.5f);
                            }

                            g.FillRectangle(new SolidBrush(resultColor), new RectangleF(p1, new SizeF(wI, Vstep / 1.5f)));
                            g.DrawRectangle(new Pen(resultColorDark, 3 * dpi), x, y, wI, Vstep / 1.5f);
                            if (actual > 0)
                            {

                                GraphicsPath path = new GraphicsPath();

                                string strTall = "";
                                if (type != BudgetValueType.Hitrate && type != BudgetValueType.SoB && type != BudgetValueType.SoM)
                                    strTall = ForkortTallBudget(actual) + " " + GetValueTypeSuffix(type);
                                else
                                    strTall = Math.Round(actual, 2).ToString() + " %";

                                path.AddString(strTall, new FontFamily("Helvetica"), (int)FontStyle.Bold, 28f * dpi, new PointF(wI + innject + offsetLeft + (5 * dpi), Vstep * i + offsetTop + (Vstep / 10)), StringFormat.GenericDefault);

                                for (int c = 1; c < 8; ++c)
                                {
                                    Pen pen3 = new Pen(Color.FromArgb(32, 255, 255, 255), c);
                                    pen3.LineJoin = LineJoin.Round;
                                    g.DrawPath(pen3, path);
                                    pen3.Dispose();
                                }

                                SolidBrush brush = new SolidBrush(Color.FromArgb(0, 0, 0));
                                g.FillPath(brush, path);
                            }
                        }
                        else // Hvis ingen flere selgere, vis tomme felt
                        {
                            // Nothing to do..
                        }
                    }

                    if (type == BudgetValueType.Hitrate || type == BudgetValueType.SoB || type == BudgetValueType.SoM)
                    {
                        float wI = maxW * ((float)list[0].target / (float)maxI);
                        g.DrawLine(new Pen(Color.Red, 10 * dpi), new PointF(wI + offsetLeft + innject, 0 + offsetTop), new PointF(wI + offsetLeft + innject, Y - offsetBottom));
                    }

                }
                catch
                {
                }
            }
            return b;
        }

        private Color Tint(Color source, Color tint, decimal alpha)
        {
            //(tint -source)*alpha + source
            int red = Convert.ToInt32(((tint.R - source.R) * alpha + source.R));
            int blue = Convert.ToInt32(((tint.B - source.B) * alpha + source.B));
            int green = Convert.ToInt32(((tint.G - source.G) * alpha + source.G));
            return Color.FromArgb(255, red, green, blue);
        }

        public string BudgetPlusMinus(decimal value, int deci = 0, bool green = false, string suffix = "")
        {
            try
            {
                string filter = "#,##0";
                if (deci == 1)
                    filter = "#,##0.0";
                else if (deci == 2)
                    filter = "#,##0.00";

                if (value == 0)
                    return main.appConfig.visningNull;
                string str = main.appConfig.visningNull;
                if (value < 0)
                    str = "<span style='color:red'>" + value.ToString(filter) + suffix + "</span>";
                if (value > 0)
                {
                    if (green)
                        str = "<span style='color:green'>+" + value.ToString(filter) + suffix + "</span>";
                    else
                        str = value.ToString(filter) + suffix;
                }

                return str;
            }
            catch
            {
                return main.appConfig.visningNull;
            }
        }

        public string ForkortTallBudget(decimal arg)
        {
            try
            {
                if (arg == 0)
                    return "0";

                var r = Math.Round((int)arg / 1000d, 1);
                if (r < 98)
                    return arg.ToString("#,##0");

                var b = Math.Round((int)arg / 1000000d, 1);
                if (b < 2)
                    return r.ToString("#,##0") + "k";

                return b.ToString() + "m";
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return "";
            }
        }


    }

    public class BudgetSelger
    {
        public string selgerkode { get; set; }
        public int timer { get; set; }
        public int dager { get; set; }
        public decimal multiplikator { get; set; }
        public decimal weight { get; set; }
    }

    public class BudgetInfo
    {
        public int budget_id { get; set; }
        public int avdeling { get; set; }
        public BudgetCategory kategori { get; set; }
        public DateTime date { get; set; }
        public int dager {get; set;}
        public decimal omsetning { get; set; }
        public decimal inntjening { get; set; }
        public decimal margin { get; set; }

        public decimal ta { get; set; }
        public BudgetValueType ta_type { get; set; }
        public decimal strom { get; set; }
        public BudgetValueType strom_type { get; set; }
        public decimal finans { get; set; }
        public BudgetValueType finans_type { get; set; }
        public decimal rtgsa { get; set; }
        public BudgetValueType rtgsa_type { get; set; }
        public decimal acc { get; set; }
        public BudgetValueType acc_type { get; set; }
        public decimal vinn { get; set; }
        public BudgetValueType vinn_type { get; set; }

        public DateTime updated { get; set; }

        public int totalt_timer { get; set; }
        public int totalt_dager { get; set; }
        public List<BudgetSelger> selgere { get; set; }
        public List<BudgetChartData> chartdata { get; set; }
        public List<BudgetChartData> chartdata_inntjen { get; set; }
        public List<BudgetChartData> chartdata_omset { get; set; }
        public List<BudgetChartData> chartdata_kvalitet { get; set; }
        public List<BudgetChartData> chartdata_effektivitet { get; set; }
        public BudgetBarChartData barchart { get; set; }
        public decimal timeElapsedCoefficient { get; set; }
        public int daysElapsed { get; set; }

        public List<BudgetCompareData> comparelist { get; set; }
        public string GetSufix(BudgetValueType type)
        {
            if (type == BudgetValueType.Antall)
                return " stk";
            else if (type == BudgetValueType.Hitrate)
                return " %";
            else if (type == BudgetValueType.Inntjening)
                return " kr";
            else if (type == BudgetValueType.Omsetning)
                return " kr";
            else if (type == BudgetValueType.SoB)
                return " %";
            else if (type == BudgetValueType.SoM)
                return " %";
            else
                return "";
        }

        public string ValueToString(decimal value, BudgetValueType type)
        {
            if (type == BudgetValueType.Antall)
                return value.ToString("#,##0") + GetSufix(type);
            else if (type == BudgetValueType.Hitrate)
                return value.ToString("#,##0.00") + GetSufix(type);
            else if (type == BudgetValueType.Inntjening)
                return value.ToString("#,##0") + GetSufix(type);
            else if (type == BudgetValueType.Omsetning)
                return value.ToString("#,##0") + GetSufix(type);
            else if (type == BudgetValueType.SoB)
                return (value * 100).ToString("#,##0.00") + GetSufix(type);
            else if (type == BudgetValueType.SoM)
                return (value * 100).ToString("#,##0.00") + GetSufix(type);
            else
                return value.ToString();
        }

        public string TypeToString(BudgetValueType type)
        {
            if (type == BudgetValueType.Poeng)
                return "Poeng";
            if (type == BudgetValueType.Antall)
                return "Antall";
            else if (type == BudgetValueType.Hitrate)
                return "Hitrate";
            else if (type == BudgetValueType.Inntjening)
                return "Inntjen";
            else if (type == BudgetValueType.Omsetning)
                return "Omset";
            else if (type == BudgetValueType.SoB)
                return "SoB";
            else if (type == BudgetValueType.SoM)
                return "SoM";
            else
                return "";
        }

        public BudgetValueType ProductToType(BudgetType product)
        {
            if (product == BudgetType.Vinnprodukt)
                return BudgetValueType.Poeng;
            else if (product == BudgetType.Strom)
                return strom_type;
            else if (product == BudgetType.TA)
                return ta_type;
            else if (product == BudgetType.Rtgsa)
                return rtgsa_type;
            else if (product == BudgetType.Finans)
                return finans_type;
            else if (product == BudgetType.Acc)
                return acc_type;
            else
                return BudgetValueType.Antall;
        }
    }

    public enum BudgetValueType { Omsetning, Inntjening, SoB, SoM, Hitrate, Antall, AntallPerDag, Empty, OmsetningPerDag, InntjeningPerDag, Poeng };

    public enum BudgetType { TA, Strom, Finans, Rtgsa, Acc, Inntjening, Omsetning, Kvalitet, Effektivitet, Vinnprodukt }

    public enum BudgetCategory { MDA, AudioVideo, SDA, Tele, Data, Cross, Kasse, Aftersales, MDASDA, None, Butikk }

    public static class BudgetCategoryClass
    {
        public static string TypeToName(BudgetCategory cat)
        {
            if (cat == BudgetCategory.MDA)
                return "MDA";
            if (cat == BudgetCategory.AudioVideo)
                return "AudioVideo";
            if (cat == BudgetCategory.SDA)
                return "SDA";
            if (cat == BudgetCategory.Tele)
                return "Tele";
            if (cat == BudgetCategory.Data)
                return "Data";
            if (cat == BudgetCategory.Cross)
                return "Cross";
            if (cat == BudgetCategory.Kasse)
                return "Kasse";
            if (cat == BudgetCategory.Aftersales)
                return "Aftersales";
            if (cat == BudgetCategory.MDASDA)
                return "MDASDA";
            if (cat == BudgetCategory.Butikk)
                return "Butikk";
            
            return "";
        }

        public static BudgetCategory NameToType(string catStr)
        {
            if (catStr == "MDA")
                return BudgetCategory.MDA;
            if (catStr == "AudioVideo")
                return BudgetCategory.AudioVideo;
            if (catStr == "SDA")
                return BudgetCategory.SDA;
            if (catStr == "Tele")
                return BudgetCategory.Tele;
            if (catStr == "Data")
                return BudgetCategory.Data;
            if (catStr == "Cross")
                return BudgetCategory.Cross;
            if (catStr == "Kasse")
                return BudgetCategory.Kasse;
            if (catStr == "Aftersales")
                return BudgetCategory.Aftersales;
            if (catStr == "MDASDA")
                return BudgetCategory.MDASDA;
            if (catStr == "Butikk")
                return BudgetCategory.Butikk;

            return BudgetCategory.None;
        }

        public static string GetSqlCategoryString(BudgetCategory cat)
        {
            if (cat == BudgetCategory.MDA)
                return "(Varegruppe >= 100 AND Varegruppe < 200) ";
            if (cat == BudgetCategory.AudioVideo)
                return "(Varegruppe >= 200 AND Varegruppe < 300) ";
            if (cat == BudgetCategory.SDA)
                return "(Varegruppe >= 300 AND Varegruppe < 400) ";
            if (cat == BudgetCategory.Tele)
                return "(Varegruppe >= 400 AND Varegruppe < 500) ";
            if (cat == BudgetCategory.Data)
                return "(Varegruppe >= 500 AND Varegruppe < 600) ";
            if (cat == BudgetCategory.Cross)
                return "(Varegruppe >= 200 AND Varegruppe < 300) OR (Varegruppe >= 400 AND Varegruppe < 500) OR (Varegruppe >= 500 AND Varegruppe < 600) ";
            if (cat == BudgetCategory.MDASDA)
                return "(Varegruppe >= 100 AND Varegruppe < 200) OR (Varegruppe >= 300 AND Varegruppe < 0) ";
            if (cat == BudgetCategory.Butikk)
                return "(Varegruppe >= 100) ";
            
            return "(Varegruppe > 0) ";
        }
    }

    public class BudgetBarChartData
    {
        public BudgetType type { get; set; }
        public decimal actual_inntjen { get; set; }
        public decimal actual_omset { get; set; }
        public decimal actual_margin { get; set; }
        public decimal target_inntjen { get; set; }
        public decimal target_omset { get; set; }
        public decimal target_margin { get; set; }
        public decimal result_inntjen { get; set; }
        public decimal result_omset { get; set; }
    }

    public class BudgetChartData
    {
        public string selgerkode { get; set; }
        public decimal target { get; set; }
        public decimal actual { get; set; }

        public BudgetChartData(string selgerkode, decimal target, decimal actual)
        {
            this.selgerkode = selgerkode;
            this.target = target;
            this.actual = actual;
        }
    }

    public class BudgetCompareData
    {
        public string type { get; set; }
        public string navn { get; set; }
        public decimal omset { get; set; }
        public decimal inntjen { get; set; }
        public decimal margin { get; set; }
        public BudgetCompareData(string type, string navn, decimal omset, decimal inntjen, decimal margin)
        {
            this.type = type;
            this.navn = navn;
            this.omset = omset;
            this.inntjen = inntjen;
            this.margin = margin;
        }
    }


}

