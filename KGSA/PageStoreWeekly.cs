using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class PageStoreWeekly : PageGenerator
    {
        public PageStoreWeekly(FormMain form, bool runInBackground, BackgroundWorker bw, System.Windows.Forms.WebBrowser webBrowser)
        {
            this.main = form;
            this.runningInBackground = runInBackground;
            this.worker = bw;
            this.browser = webBrowser;
        }

        public bool BuildPage_Report(string strCat, string strHash, string htmlPage, DateTime date)
        {
            pickedDate = date;
            bool abort = main.HarSisteVersjonStore(strCat, strHash);
            try
            {
                if (!runningInBackground && !abort) main.timewatch.Start();
                if (!runningInBackground) main.appConfig.savedStorePage = strCat;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + strCat + "]..");
                    OpenPage_Loading();

                    doc = new List<string>();

                    main.openXml.DeleteDocument(strCat, pickedDate);

                    AddPage_Start(true, "Lager (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");
                    AddPage_Title("Lager (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");

                    ShowProgress();

                    MakeJumplist();

                    ShowProgress();

                    MakeWeeklyReport(pickedDate);

                    AddPage_End();

                    if (FormMain.stopRanking)
                    {
                        main.ClearHash(strCat);
                        Log.n("Lasting avbrutt", Color.Red);
                        OpenPage_Stopped();
                        FormMain.stopRanking = false;
                    }
                    else
                    {
                        File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                        OpenPage(htmlPage);
                        if (!runningInBackground)
                            Log.n("Side [" + strCat + "] tok " + main.timewatch.Stop() + " sekunder.", Color.Black, true);
                    }
                }
                else
                    OpenPage(htmlPage);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!runningInBackground)
                {
                    OpenPage_Error();
                    FormError errorMsg = new FormError("Feil ved generering av side for [" + strCat + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
            return false;
        }

        public bool BuildPage_Overview(string strCat, string strHash, string htmlPage, DateTime date)
        {
            pickedDate = date;
            bool abort = main.HarSisteVersjonStore(strCat, strHash);
            try
            {
                if (!runningInBackground && !abort) main.timewatch.Start();
                if (!runningInBackground) main.appConfig.savedStorePage = strCat;
                if (!abort)
                {
                    Log.n("Oppdaterer [" + strCat + "]..");
                    OpenPage_Loading();

                    doc = new List<string>();

                    main.openXml.DeleteDocument(strCat, pickedDate);

                    AddPage_Start(true, "Lager (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");
                    AddPage_Title("Lager (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");

                    ShowProgress();

                    MakeOverview();

                    ShowProgress();

                    DateTime latest = main.database.tableWeekly.GetLatestDate(main.appConfig.Avdeling);

                    MakeWeeklyReport(latest);

                    AddPage_End();

                    if (FormMain.stopRanking)
                    {
                        main.ClearHash(strCat);
                        Log.n("Lasting avbrutt", Color.Red);
                        OpenPage_Stopped();
                        FormMain.stopRanking = false;
                    }
                    else
                    {
                        File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                        OpenPage(htmlPage);
                        if (!runningInBackground)
                            Log.n("Side [" + strCat + "] tok " + main.timewatch.Stop() + " sekunder.", Color.Black, true);
                    }
                }
                else
                    OpenPage(htmlPage);
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                if (!runningInBackground)
                {
                    OpenPage_Error();
                    FormError errorMsg = new FormError("Feil ved generering av side for [" + strCat + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
            return false;
        }

        private void MakeWeeklyReport(DateTime date)
        {
            try
            {
                Log.d("Henter ukeannonse..");
                DataTable table = main.obsolete.GetWeekly(date);
                if (table == null || table.Rows.Count == 0)
                {
                    AddWarning("Ingen oppføringer funnet på angitt dato: " + date.ToString("dddd d. MMMM yyyy", FormMain.norway));
                    return;
                }

                for (int d = 1; d < 6; d++)
                {
                    if (FormMain.stopRanking)
                        return;

                    var rows = table.Select("Kategori = '" + d + "'");
                    DataTable tableCategory = rows.Any() ? rows.CopyToDataTable() : table.Clone();

                    string category = "";
                    if (d == 1)
                        category = "MDA";
                    else if (d == 2)
                        category = "AudioVideo";
                    else if (d == 3)
                        category = "SDA";
                    else if (d == 4)
                        category = "Tele";
                    else if (d == 5)
                        category = "Data";

                    main.openXml.SaveDocument(tableCategory, "LagerUkeAnnonser", category, pickedDate,
                        "Ukeannonser lagerstatus - " + date.ToString("dddd d. MMMM yyyy", FormMain.norway) + " - " + category
                        + " - Uke " + main.database.GetIso8601WeekOfYear(date));

                    AddTable_Start(category + " - " + date.ToString("dddd d. MMMM yyyy", FormMain.norway)
                        + " - Uke " + main.database.GetIso8601WeekOfYear(date));

                    AddTable_Header_Start();
                    AddTable_Header_Name("Varekode", 120, "", Sorter_Type_Text);

                    AddTable_Header_Name("Varetekst", 250, "", Sorter_Type_Text);
                    AddTable_Header_Name("Merke", 120, "", Sorter_Type_Text);
                    AddTable_Header_Name("Pris", 80, "", Sorter_Type_Digit);
                    AddTable_Header_Name("Lager", 70, "", Sorter_Type_Digit);
                    AddTable_Header_Name("Nettlager", 70, "", Sorter_Type_Text);

                    AddTable_Header_End();
                    AddTable_Body_Start();

                    for (int i = 0; i < tableCategory.Rows.Count; i++)
                    {
                        int stock = Convert.ToInt32(tableCategory.Rows[i][TableWeekly.KEY_PRODUCT_STOCK]);
                        int stockInternet = Convert.ToInt32(tableCategory.Rows[i][TableWeekly.KEY_PRODUCT_STOCK_INTERNET]);

                        AddTable_Row_Start();

                        AddTable_Row_Cell("<a href='#linkmv" + tableCategory.Rows[i]["ProductCode"] + "'>" + tableCategory.Rows[i]["ProductCode"] + "</a>",
                            "", Class_Style_Text_Cat);
                        AddTable_Row_Cell(main.tools.TextStyle_Shorten(tableCategory.Rows[i]["Varetekst"].ToString(), 30),
                            "", Class_Style_Text_Cat);
                        AddTable_Row_Cell(main.tools.TextStyle_Shorten(tableCategory.Rows[i]["MerkeNavn"].ToString(), 17), 
                            "", Class_Style_Text_Cat);
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(Convert.ToDecimal(tableCategory.Rows[i][TableWeekly.KEY_PRODUCT_PRIZE_INTERNET])),
                            "", Class_Style_Generic);

                        if (stock > 0)
                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(stock),
                                "background-color: #71ba51 !important;", Class_Style_Small);
                        else
                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(stock),
                                "background-color: #e98263 !important;", Class_Style_Small);

                        if (stockInternet > 0)
                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(stockInternet) + "+",
                                "background-color: #71ba51 !important;", Class_Style_Small);
                        else
                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(stockInternet),
                                "background-color: #e98263 !important;", Class_Style_Small);

                        AddTable_Row_End();
                    }

                    AddTable_Body_End();

                    AddTable_End();
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil ved generering av side: " + ex.Message, Color.Red);
            }
        }

        public void MakeOverview()
        {
            try
            {
                Log.d("Henter ukeannonse oversikt..");
                
                DataTable table = main.database.tableWeekly.GetWeeklyList(main.appConfig.Avdeling);
                if (table == null || table.Rows.Count == 0)
                {
                    AddWarning("Fant ingen ukeannonse oppføringer");
                    return;
                }

                main.openXml.SaveDocument(table, "LagerUkeAnnonserOversikt", "Ukeannonser", pickedDate,
                    "Ukeannonser lagerstatus - " + pickedDate.ToString("dddd d. MMMM yyyy", FormMain.norway) + " - Liste");

                int week = main.database.GetIso8601WeekOfYear(Convert.ToDateTime(table.Rows[0][TableWeekly.KEY_DATE]));
                bool weekSwitch = true;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    DateTime listDate = Convert.ToDateTime(table.Rows[i]["Date"]);

                    if (weekSwitch)
                    {
                        AddTable_Start("Lagerstatus for Ukeannonser - Uke " + main.database.GetIso8601WeekOfYear(listDate) + "");
                        AddTable_Header_Start();
                        AddTable_Header_Name("Ukedag", 100, "", Sorter_Type_Text);
                        AddTable_Header_Name("Totalt for butikk", 120, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_Name("Nettbutikk", 90, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_Name("MDA/SDA", 90, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_Name("Tele", 90, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_Name("AudioVideo", 90, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_Name("Computer", 90, "", Sorter_Type_Digit);
                        AddTable_Header_Name("%", 40, "", Sorter_Type_Procent);
                        AddTable_Header_End();
                        AddTable_Body_Start();
                        weekSwitch = false;
                    }

                    int currentWeek = main.database.GetIso8601WeekOfYear(listDate);
                    int listNumberOfProducts = Convert.ToInt32(table.Rows[i]["NumberOfProducts"]);
                    int listNoInStock = Convert.ToInt32(table.Rows[i]["NoInStock"]);
                    int listNoInInetStock = Convert.ToInt32(table.Rows[i]["NoInInetStock"]);

                    int listInStockMdaSda = Convert.ToInt32(table.Rows[i]["InStockMdaSda"]);
                    int listTotalMdaSda = Convert.ToInt32(table.Rows[i]["TotalMdaSda"]);
                    int listInStockTelecom = Convert.ToInt32(table.Rows[i]["InStockTelecom"]);
                    int listTotalTelecom = Convert.ToInt32(table.Rows[i]["TotalTelecom"]);
                    int listInStockAudioVideo = Convert.ToInt32(table.Rows[i]["InStockAudioVideo"]);
                    int listTotalAudioVideo = Convert.ToInt32(table.Rows[i]["TotalAudioVideo"]);
                    int listInStockComputer = Convert.ToInt32(table.Rows[i]["InStockComputer"]);
                    int listTotalComputer = Convert.ToInt32(table.Rows[i]["TotalComputer"]);

                    AddTable_Row_Start();

                    AddTable_Row_Cell("<a href='#ukenytt=" + listDate.ToString("dd.MM.yyyy", FormMain.norway) + "'>"
                        + listDate.ToString("dddd .dd", FormMain.norway).ToUpper() + "</a>", "", Class_Style_Text_Cat);

                    AddTable_Row_Cell(listNoInStock + " av " + listNumberOfProducts, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listNoInStock, listNumberOfProducts, false, false), "", Class_Style_Small);

                    AddTable_Row_Cell(listNoInInetStock + " av " + listNumberOfProducts, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listNoInInetStock, listNumberOfProducts, false, false), "", Class_Style_Small);

                    AddTable_Row_Cell(listInStockMdaSda + " av " + listTotalMdaSda, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listInStockMdaSda, listTotalMdaSda, false, false), "", Class_Style_Small);

                    AddTable_Row_Cell(listInStockTelecom + " av " + listTotalTelecom, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listInStockTelecom, listTotalTelecom, false, false), "", Class_Style_Small);

                    AddTable_Row_Cell(listInStockAudioVideo + " av " + listTotalAudioVideo, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listInStockAudioVideo, listTotalAudioVideo, false, false), "", Class_Style_Small);

                    AddTable_Row_Cell(listInStockComputer + " av " + listTotalComputer, "", Class_Style_Small);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listInStockComputer, listTotalComputer, false, false), "", Class_Style_Small);

                    AddTable_Row_End();

                    if (table.Rows.Count > i + 1)
                    {
                        week = main.database.GetIso8601WeekOfYear(Convert.ToDateTime(table.Rows[i + 1][TableWeekly.KEY_DATE]));
                        if (week != currentWeek)
                        {
                            AddTable_End();
                            week = currentWeek;
                            weekSwitch = true;
                        }
                    }
                    else
                        AddTable_End();
                }
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil ved generering av side: " + ex.Message, Color.Red);
            }
        }

        public void MakeJumplist()
        {
            try
            {
                Log.d("Henter ukenytt liste..");
                DataTable table = main.database.tableWeekly.GetWeeklyList(main.appConfig.Avdeling);

                doc.Add("<form class='hidePdf' name='jumpSelect'><br>");
                doc.Add("<select name='menu' onChange='window.document.location.href=this.options[this.selectedIndex].value;' value='GO'>");

                doc.Add("<option value='#ukenytt=" + pickedDate.ToString("dd.MM.yyyy", FormMain.norway) + "'>Velg status oppdatering</option>");
                doc.Add("<option value='#ukenytt=list'>Oversikt</option>");
                foreach (DataRow dRow in table.Rows)
                {

                    DateTime date = Convert.ToDateTime(dRow["Date"]);
                    string selectedStr = "";
                    if (pickedDate.Date == date.Date)
                        selectedStr = " selected";

                    doc.Add("<option value='#ukenytt=" + date.ToString("dd.MM.yyyy", FormMain.norway) + "' " + selectedStr + ">"
                        + date.ToShortDateString() + " - " + date.ToString("dddd") + "</option>");
                }

                doc.Add("</select>");
                doc.Add("</form>");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.n("Feil ved generering av side: " + ex.Message, Color.Red);
            }
        }
    }
}
