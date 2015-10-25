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
    public class PageStorePrisguide : PageGenerator
    {
        public PageStorePrisguide(FormMain form, bool runInBackground, BackgroundWorker bw, System.Windows.Forms.WebBrowser webBrowser)
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
                if (!runningInBackground) main.savedStorePage = strCat;
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
                if (!runningInBackground) main.savedStorePage = strCat;
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

                    DateTime latest = main.database.tablePrisguide.GetLatestDate(main.appConfig.Avdeling);

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
                DataTable table = main.obsolete.GetPopularPrisguideProducts(date);
                if (table == null || table.Rows.Count == 0)
                {
                    AddWarning("Ingen oppføringer funnet på angitt dato: " + date.ToString("dddd d. MMMM yyyy", FormMain.norway));
                    return;
                }

                main.openXml.SaveDocument(table, "LagerPrisguide", "Prisguide.no", date,
                    "De mest populære produkter på Prisguide.no - " + date.ToString("dddd d. MMMM yyyy", FormMain.norway));

                AddTable_Start("De mest populære produktene på Prisguide.no - " + date.ToString("dddd d. MMMM yyyy", FormMain.norway));

                AddTable_Header_Start();
                AddTable_Header_Name("#", 35, "", Sorter_Type_Text);
                AddTable_Header_Name("Varekoder", 120, "", Sorter_Type_Text);
                AddTable_Header_Name("Varetekst", 220, "", Sorter_Type_Text);
                AddTable_Header_Name("Merke", 120, "", Sorter_Type_Text);
                AddTable_Header_Name("Pris", 80, "", Sorter_Type_Digit);
                AddTable_Header_Name("Lager", 70, "", Sorter_Type_Digit);
                AddTable_Header_Name("Nettlager", 70, "", Sorter_Type_Text);
                AddTable_Header_End();
                AddTable_Body_Start();

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int status = Convert.ToInt32(table.Rows[i][TablePrisguide.KEY_PRISGUIDE_STATUS]);
                    int stock = Convert.ToInt32(table.Rows[i][TablePrisguide.KEY_PRODUCT_STOCK]);
                    int stockInternet = Convert.ToInt32(table.Rows[i][TablePrisguide.KEY_PRODUCT_STOCK_INTERNET]);
                    string tekst = table.Rows[i]["Varetekst"].ToString();

                    AddTable_Row_Start();

                    AddTable_Row_Cell(table.Rows[i]["Position"].ToString(), "", Class_Style_Small);

                    if (table.Rows[i]["ProductCode"].Equals(""))
                        AddTable_Row_Cell("<a href='#external=http://www.prisguide.no/produkt/" + table.Rows[i]["PrisguideId"] + "'>Link</a>",
                            "", Class_Style_Text_Cat);
                    else
                        AddTable_Row_Cell("<a href='#external=http://www.prisguide.no/produkt/" + table.Rows[i]["PrisguideId"] + "'>" + table.Rows[i]["ProductCode"] + "</a>",
                            "", Class_Style_Text_Cat);

                    if (status == 0 && !String.IsNullOrEmpty(tekst))
                        AddTable_Row_Cell(main.tools.TextStyle_Shorten(tekst, 30), "", Class_Style_Text_Cat);
                    else if (status == 0 && String.IsNullOrEmpty(tekst))
                        AddTable_Row_Cell(main.tools.TextStyle_Shorten(tekst, 30), "color:#454545;text-align: center;", Class_Style_Text_Cat);
                    else
                        AddTable_Row_Cell(PrisguideProduct.GetStatusStatic(status), "color:#454545;text-align: center;", Class_Style_Text_Cat);

                    AddTable_Row_Cell(main.tools.TextStyle_Shorten(table.Rows[i]["MerkeNavn"].ToString(), 17),
                        "", Class_Style_Text_Cat);

                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(table.Rows[i][TablePrisguide.KEY_PRODUCT_PRIZE_INTERNET]), "", Class_Style_Generic);

                    if (stock > 0)
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(stock), "background-color: #71ba51 !important;", Class_Style_Small);
                    else if (status == 0 && stock == 0)
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(stock), "background-color: #e98263 !important;", Class_Style_Small);
                    else
                        AddTable_Row_Cell("&nbsp;", "", Class_Style_Small);

                    if (stockInternet > 0 && status == 0)
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(stockInternet), "background-color: #71ba51 !important;", Class_Style_Small);
                    else if (stockInternet == 0 && status == 0)
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(stockInternet), "background-color: #e98263 !important;", Class_Style_Small);
                    else
                        doc.Add("<td class='numbers-small'>&nbsp;</td>");

                    AddTable_Row_End();
                }

                AddTable_Body_End();

                AddTable_End();
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
                
                DataTable table = main.database.tablePrisguide.GetPrisguideList(main.appConfig.Avdeling);
                if (table == null || table.Rows.Count == 0)
                {
                    AddWarning("Fant ingen prisguide oppføringer");
                    return;
                }

                main.openXml.SaveDocument(table, "LagerPrisguideOversikt", "Prisguide.no", pickedDate,
                    "De mest populære produkter på Prisguide.no - " + pickedDate.ToString("dddd d. MMMM yyyy", FormMain.norway));

                AddTable_Start("Lagerstatus for Prisguide varer - " + DateTime.Now.ToString("dddd d. MMMM yyyy", FormMain.norway));

                AddTable_Header_Start();
                AddTable_Header_Name("Dato", 120, "", Sorter_Type_Text);
                AddTable_Header_Name("Antall varekoder", 100, "", Sorter_Type_Text);
                AddTable_Header_Name("Butikk", 40, "", Sorter_Type_Text);
                AddTable_Header_Name("%", 60, "", Sorter_Type_Procent);
                AddTable_Header_Name("Nettlager", 100, "", Sorter_Type_Text);
                AddTable_Header_Name("%", 60, "", Sorter_Type_Procent);
                AddTable_Header_End();
                AddTable_Body_Start();

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    DateTime listDate = Convert.ToDateTime(table.Rows[i]["Date"]);
                    int listTotal = Convert.ToInt32(table.Rows[i]["Total"]);
                    int listNotInStock = Convert.ToInt32(table.Rows[i]["NoStock"]);
                    int listNotInInetStock = Convert.ToInt32(table.Rows[i]["NoInetStock"]);

                    AddTable_Row_Start();

                    AddTable_Row_Cell("<a href='#prisguide=" + listDate.ToString("dd.MM.yyyy", FormMain.norway) + "'>"
                        + listDate.ToString("dddd d. MMMM yyyy", FormMain.norway) + "</a>", "", Class_Style_Text_Cat);
                    
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(listTotal), "", Class_Style_Small);
                    AddTable_Row_Cell((listTotal - listNotInStock) + " av " + listTotal, "", Class_Style_Small);

                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listTotal - listNotInStock, listTotal, false, false), "", Class_Style_Small);
                    AddTable_Row_Cell((listTotal - listNotInInetStock) + " av " + listTotal, "", Class_Style_Small);

                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(listTotal - listNotInInetStock, listTotal, false, false), "", Class_Style_Small);

                    AddTable_Row_End();
                }

                AddTable_Body_End();
                AddTable_End();
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
                Log.d("Henter prisguide liste..");
                DataTable table = main.database.tablePrisguide.GetPrisguideList(main.appConfig.Avdeling);

                doc.Add("<form class='hidePdf' name='jumpSelect'><br>");
                doc.Add("<select name='menu' onChange='window.document.location.href=this.options[this.selectedIndex].value;' value='GO'>");

                doc.Add("<option value='#prisguide=" + pickedDate.ToString("dd.MM.yyyy", FormMain.norway) + "'>Velg prisguide oppdatering</option>");
                doc.Add("<option value='#prisguide=list'>Oversikt</option>");
                foreach (DataRow dRow in table.Rows)
                {
                    DateTime date = Convert.ToDateTime(dRow["Date"]);
                    string selectedStr = "";
                    if (pickedDate.Date == date.Date)
                        selectedStr = " selected";

                    doc.Add("<option value='#prisguide=" + date.ToString("dd.MM.yyyy", FormMain.norway) + "' " + selectedStr + ">"
                        + date.ToString("dddd d. MMMM yyyy", FormMain.norway) + " - (" + Convert.ToInt32(dRow["NoStock"]) + " / " + Convert.ToInt32(dRow["Total"]) + "</option>");
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
