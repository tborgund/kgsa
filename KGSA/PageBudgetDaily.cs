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
    public class PageBudgetDaily : PageRanking
    {
        public PageBudgetDaily(FormMain form, bool runInBackground, BackgroundWorker bw, System.Windows.Forms.WebBrowser webBrowser)
        {
            this.main = form;
            this.runningInBackground = runInBackground;
            this.worker = bw;
            this.browser = webBrowser;
        }

        public bool BuildPage(BudgetCategory cat, string strHash, string htmlPage, DateTime date)
        {
            pickedDate = date;
            string katArg = BudgetCategoryClass.TypeToName(cat);
            bool abort = main.HarSisteVersjonBudget(cat, strHash);
            try
            {
                if (!runningInBackground && !abort) main.timewatch.Start();
                if (!runningInBackground)
                    main.savedBudgetPage = cat;
                if (!abort)
                {
                    Logg.Log("Oppdaterer [" + BudgetCategoryClass.TypeToName(cat) + "]..");
                    if (!runningInBackground) browser.Navigate(FormMain.htmlLoading);

                    doc = new List<string>();

                    main.openXml.DeleteDocument(BudgetCategoryClass.TypeToName(cat), pickedDate);

                    AddPage_Start(true, "Ranking: " + BudgetCategoryClass.TypeToName(cat));
                    AddPage_Title("Budsjett (" + main.avdeling.Get(main.appConfig.Avdeling) + ")");

                    ShowProgress();

                    BudgetImporter importer = new BudgetImporter(main, DateTime.Now);

                    if (main.tableMacroQuick == null)
                        main.tableMacroQuick = importer.ImportElguideBudget(main.appConfig.Avdeling);

                    if (main.appConfig.dailyBudgetIncludeInQuickRanking)
                        MakeDailyBudgetFromDatabase();

                    DailyBudgetMacroInfo budgetInfo = null;
                    if (main.appConfig.macroImportQuickSales)
                        budgetInfo = importer.ImportElguideServices();

                    MakeDailyBudgetFromElguide(budgetInfo);

                    doc.Add(Resources.htmlEnd);
                    if (FormMain.stopRanking)
                    {
                        main.ClearHash(katArg);
                        Logg.Log("Ranking stoppet.", Color.Red);
                        browser.Navigate(FormMain.htmlStopped);
                        FormMain.stopRanking = false;
                    }
                    else
                    {
                        File.WriteAllLines(htmlPage, doc.ToArray(), Encoding.Unicode);
                        if (!runningInBackground)
                            browser.Navigate(htmlPage);
                        if (!runningInBackground)
                            Logg.Log("Ranking [" + katArg + "] tok " + main.timewatch.Stop() + " sekunder.", Color.Black, true);
                    }
                }
                else if (!runningInBackground)
                    browser.Navigate(htmlPage);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                if (!runningInBackground)
                {
                    browser.Navigate(FormMain.htmlError);
                    FormError errorMsg = new FormError("Feil ved generering av ranking for [" + katArg + "]", ex);
                    errorMsg.ShowDialog();
                }
            }
            return false;
        }

        private void MakeDailyBudgetFromDatabase()
        {
            try
            {
                Logg.Debug("Henter dagens budsjett fra database..");
                DataTable tableBudget = main.database.tableDailyBudget.GetBudgetFromDate(main.appConfig.Avdeling, DateTime.Now);

                if (tableBudget == null || tableBudget.Rows.Count < 6 || main.tableMacroQuick == null || main.tableMacroQuick.Rows.Count < 6)
                {
                    AddWarning("Mangler data for budsjett");
                    return;
                }

                AddTable_Start("Resultat mot budsjett " + main.avdeling.Get(main.appConfig.Avdeling)
                    + " - " + DateTime.Now.ToString("dddd d. MMMM yyyy  HH:mm", FormMain.norway) + "");

                AddTable_Header_Start();
                AddTable_Header_Name("Category", 95, "background:#bfd2e2;", Sorter_Type_Text);

                AddTable_Header_Name("Sales", 60, "background:#bfd2e2;border-left:2px solid #000;", Sorter_Type_Digit);
                AddTable_Header_Name("Budget", 60, "background:#bfd2e2;", Sorter_Type_Digit);
                AddTable_Header_Name("Index", 60, "background:#bfd2e2;", Sorter_Type_Procent);

                AddTable_Header_Name("GM %", 60, "background:#bfd2e2;border-left:2px solid #000;", Sorter_Type_Digit);
                AddTable_Header_Name("Budget", 60, "background:#bfd2e2;", Sorter_Type_Digit);
                AddTable_Header_Name("Variance", 60, "background:#bfd2e2;", Sorter_Type_Procent);

                AddTable_Header_Name("GM", 60, "background:#bfd2e2;border-left:2px solid #000;", Sorter_Type_Digit);
                AddTable_Header_Name("Budget", 60, "background:#bfd2e2;", Sorter_Type_Digit);
                AddTable_Header_Name("Variance", 60, "background:#bfd2e2;", Sorter_Type_Procent);
                AddTable_Header_End();

                AddTable_Body_Start();

                for (int i = 0; i < main.tableMacroQuick.Rows.Count; i++)
                {
                    if (Convert.ToInt32(main.tableMacroQuick.Rows[i]["Favoritt"]) != main.appConfig.Avdeling)
                        break;

                    string kategori = main.tableMacroQuick.Rows[i]["Avdeling"].ToString();
                    decimal omset_budget = 0, inntjen_budget = 0, margin_budget = 0;
                    decimal omset = Convert.ToDecimal(main.tableMacroQuick.Rows[i]["Omsetn"]);
                    decimal inntjen = Convert.ToDecimal(main.tableMacroQuick.Rows[i]["Fortjeneste"]);
                    decimal margin = Convert.ToDecimal(main.tableMacroQuick.Rows[i]["Margin"]);
                    foreach (DataRow row in tableBudget.Rows)
                    {
                        if (kategori.ToLower().Equals(row[TableDailyBudget.KEY_BUDGET_TYPE].ToString().ToLower()))
                        {
                            omset_budget = Convert.ToDecimal(row[TableDailyBudget.KEY_BUDGET_SALES]);
                            inntjen_budget = Convert.ToDecimal(row[TableDailyBudget.KEY_BUDGET_GM]);
                            margin_budget = Convert.ToDecimal(row[TableDailyBudget.KEY_BUDGET_GM_PERCENT]);
                            break;
                        }
                    }

                    if (kategori.ToLower().Equals("total")) // siste row
                    {
                        AddTable_Body_End();
                        AddTable_Footer_Start();
                    }

                    AddTable_Row_Start();

                    AddTable_Row_Cell(kategori, "", Class_Style_Numbers_Text_Cat);

                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(omset), "border-left:2px solid #000;text-align: right;", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(omset_budget), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Percent(omset, omset_budget, true, true), "", Class_Style_Numbers_Gen);

                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(margin, 1, "%"), "border-left:2px solid #000;text-align: right;", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(margin_budget, 1, "%"), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(margin - margin_budget, 1, "%", true, true), "", Class_Style_Numbers_Gen);

                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(inntjen), "border-left:2px solid #000;text-align: right;", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(inntjen_budget), "", Class_Style_Numbers_Gen);
                    AddTable_Row_Cell(main.tools.NumberStyle_Normal(inntjen - inntjen_budget, 0, "", true, true), "", Class_Style_Numbers_Gen);

                    AddTable_Row_End();

                    if (kategori.ToLower().Equals("total")) // siste row
                        AddTable_Footer_End();
                }

                AddTable_End();
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Feil ved generering av side: " + ex.Message, Color.Red);
            }
            return;
        }

        private void MakeDailyBudgetFromElguide(DailyBudgetMacroInfo budgetInfo)
        {
            try
            {
                if (main.tableMacroQuick == null || main.tableMacroQuick.Rows.Count < 5)
                {
                    AddWarning("Mangler tall fra Elguide. Eksporter på nytt fra meny 136 i Elguide og prøv igjen");
                    return;
                }

                for (int f = 0; f < main.tableMacroQuick.Rows.Count; f++)
                {
                    int intAvd = Convert.ToInt32(main.tableMacroQuick.Rows[f][0]);

                    AddTable_Start("Dagens tall for " + main.avdeling.Get(Convert.ToInt32(main.tableMacroQuick.Rows[f]["Favoritt"].ToString()))
                        + " - " + DateTime.Now.ToString("dddd d. MMMM yyyy  HH:mm", FormMain.norway) + ")");

                    AddTable_Header_Start();
                    AddTable_Header_Name("Category", 95, "", Sorter_Type_Text);

                    AddTable_Header_Name("Salg", 60, "", Sorter_Type_Text);
                    AddTable_Header_Name("Omset", 60, "", Sorter_Type_Digit);
                    AddTable_Header_Name("Fritt", 60, "", Sorter_Type_Digit);

                    AddTable_Header_Name("Inntjen.", 60, "", Sorter_Type_Digit);
                    AddTable_Header_Name("Margin", 60, "", Sorter_Type_Procent);
                    AddTable_Header_Name("Rabatt", 60, "", Sorter_Type_Digit);

                    if (main.appConfig.macroImportQuickSales && intAvd == main.appConfig.Avdeling
                        && budgetInfo != null && budgetInfo.Salg != null)
                    {
                        if (budgetInfo.Salg.Count > 0)
                        {
                            AddTable_Header_Name("RTG/SA", 60, "background:#80c34a;", Sorter_Type_Digit);
                            AddTable_Header_Name("Omset.", 60, "background:#80c34a;", Sorter_Type_Digit);
                            AddTable_Header_Name("Inntjen.", 60, "background:#80c34a;", Sorter_Type_Digit);
                            AddTable_Header_Name("%", 60, "background:#80c34a;", Sorter_Type_Digit, "Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer");
                        }
                    }

                    AddTable_Header_End();

                    AddTable_Body_Start();

                    double KgsaAntallTot = 0, KgsaBtokrTot = 0, KgsaSalgsprisTot = 0;
                    for (int i = f; i < main.tableMacroQuick.Rows.Count; i++)
                    {
                        if (intAvd != (Convert.ToInt32(main.tableMacroQuick.Rows[i][0])))
                        {
                            f--;
                            break;
                        }
                        f++;

                        if (main.tableMacroQuick.Rows.Count == i + 1 || main.tableMacroQuick.Rows[i]["Avdeling"].ToString() == "TOTALT") // siste row
                        {
                            AddTable_Body_End();
                            AddTable_Footer_Start();
                        }

                        AddTable_Row_Cell(main.tableMacroQuick.Rows[i]["Avdeling"].ToString(), "", Class_Style_Numbers_Text_Cat);

                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(main.tableMacroQuick.Rows[i]["Salg"], 0, "", true), "", Class_Style_Numbers_Small);
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(main.tableMacroQuick.Rows[i]["Omsetn"], 0, "", true), "", Class_Style_Numbers_Gen);
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(main.tableMacroQuick.Rows[i]["Fritt"], 0, "", true), "", Class_Style_Numbers_Gen);

                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(main.tableMacroQuick.Rows[i]["Fortjeneste"], 0, "", true), "", Class_Style_Numbers_Gen);
                        AddTable_Row_Cell(main.tools.NumberStyle_Percent(main.tableMacroQuick.Rows[i]["Margin"], 100, true, false, 0), "", Class_Style_Numbers_Percent);
                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(main.tableMacroQuick.Rows[i]["Rabatt"], 0, "", true), "", Class_Style_Numbers_Gen);

                        if (main.appConfig.macroImportQuickSales && intAvd == main.appConfig.Avdeling
                            && budgetInfo != null && budgetInfo.Salg != null)
                        {
                            try
                            { 
                                if (budgetInfo.Salg.Count > 0)
                                {
                                    if ("TOTALT" == main.tableMacroQuick.Rows[i]["Avdeling"].ToString())
                                    {
                                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(KgsaAntallTot, 0), "", Class_Style_Numbers_Small);
                                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(KgsaSalgsprisTot, 0), "", Class_Style_Numbers_Gen);
                                        AddTable_Row_Cell(main.tools.NumberStyle_Normal(KgsaBtokrTot, 0), "", Class_Style_Numbers_Gen);
                                        AddTable_Row_Cell(main.tools.NumberStyle_Percent(KgsaBtokrTot, main.tableMacroQuick.Rows[i]["Fortjeneste"]), "", Class_Style_Numbers_Percent);
                                    }
                                    else
                                    {
                                        bool found = false;
                                        foreach (DailyBudgetMacroInfoItem item in budgetInfo.Salg)
                                        {
                                            if (item.Type == main.tableMacroQuick.Rows[i]["Avdeling"].ToString())
                                            {
                                                AddTable_Row_Cell(main.tools.NumberStyle_Normal(item.Antall, 0), "", Class_Style_Numbers_Small);
                                                KgsaAntallTot += item.Antall;
                                                AddTable_Row_Cell(main.tools.NumberStyle_Normal(item.Salgspris, 0), "", Class_Style_Numbers_Gen);
                                                KgsaSalgsprisTot += item.Salgspris;
                                                AddTable_Row_Cell(main.tools.NumberStyle_Normal(item.Btokr, 0), "", Class_Style_Numbers_Gen);
                                                KgsaBtokrTot += item.Btokr;
                                                AddTable_Row_Cell(main.tools.NumberStyle_Percent(item.Btokr / Convert.ToDouble(main.tableMacroQuick.Rows[i]["Fortjeneste"]), 
                                                    100, true, false, 0), "", Class_Style_Numbers_Percent);

                                                found = true;
                                                break;
                                            }
                                        }
                                        if (!found)
                                        {
                                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(0, 0), "", Class_Style_Numbers_Small);
                                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(0, 0), "", Class_Style_Numbers_Gen);
                                            AddTable_Row_Cell(main.tools.NumberStyle_Normal(0, 0), "", Class_Style_Numbers_Gen);
                                            AddTable_Row_Cell(main.tools.NumberStyle_Percent(0, 100), "", Class_Style_Numbers_Percent);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Logg.Unhandled(ex);
                                Logg.Log("Unntak ved skriving av RTGSA salg: " + ex.Message, Color.Red);
                            }
                        }

                        AddTable_Row_End();

                        if (main.tableMacroQuick.Rows.Count == i + 1 || main.tableMacroQuick.Rows[i]["Avdeling"].ToString() == "TOTALT") // siste row
                            AddTable_Footer_End();

                    }

                    AddTable_End();
                }
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Uventet feil oppstod under generering av daglig budsjett tabellen", Color.Red);
            }
        }
    }
}
