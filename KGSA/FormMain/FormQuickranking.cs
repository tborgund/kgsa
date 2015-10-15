using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    partial class FormMain
    {
        private KveldstallInfo kvdelstallInfo;

        public void startDelayedAutoQuickImport()
        {
            if (IsBusy())
                return;

            processing.SetVisible = true;
            processing.SetProgressStyle = ProgressBarStyle.Continuous;
            processing.SetBackgroundWorker = bwQuickAuto;
            for (int b = 0; b < 100; b++)
            {
                processing.SetText = "Starter automatisk kvelds ranking om " + (((b / 10) * -1) + 10) + " sekunder..";
                processing.SetValue = b;
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);

                if (processing.userPushedCancelButton)
                {
                    Logg.Log("Brukeren avbrøt handlingen.");
                    return;
                }
            }

            RestoreWindow();

            processing.SetProgressStyle = ProgressBarStyle.Marquee;
            processing.SetBackgroundWorker = bwQuickAuto;
            processing.SetText = "Kjører kveldsranking rutine..";
            Logg.Log("Kjører kveldsranking rutine..");
            bwQuickAuto.RunWorkerAsync();

        }

        private void bwQuickAuto_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgressStart();
            autoMode = true;

            var macroAttempt = 0;
            var macroMaxAttempts = 4;

            retrymacro:
            macroAttempt++;

            FormMacro formMacro = (FormMacro)StartMacro(dbTilDT, macroProgramQuick, bwQuickAuto, macroAttempt);
            e.Result = formMacro.errorCode;
            kvdelstallInfo = formMacro.KveldstallInfo;

            if (formMacro.errorCode == 6)
            {
                Logg.Log("En kritisk feil forhindret makro i å utføre sine oppgaver. Se logg for detaljer.", Color.Red);
                e.Result = 6;
                return;
            }

            if (formMacro.tableQuick != null)
            {
                if (formMacro.tableQuick.Rows.Count < 5 || formMacro.errorCode != 0)
                {
                    System.Threading.Thread.Sleep(3000);
                    if (macroAttempt < macroMaxAttempts && formMacro.errorCode != 2) // Vi har flere forsøk igjen, samt bruker har ikke avbrutt prosessen
                        goto retrymacro;
                    if (formMacro.errorCode != 2)
                        Logg.Log("Kveldstall: Feil oppstod under kveldsranking automatisering. Feilbeskjed: " + formMacro.errorMessage + " Kode: " + formMacro.errorCode, Color.Red);
                    return;
                }
            }


            if (appConfig.dailyBudgetIncludeInQuickRanking)
            {
                BudgetImporter budgetImporter = new BudgetImporter(this);
                budgetImporter.FindAndDownloadBudget(bwQuickAuto);
            }

            MakeQuickHtml(formMacro.tableQuick);

            if (appConfig.quickInclService && appConfig.AutoService && service.dbServiceDatoFra.Date != service.dbServiceDatoTil.Date)
            {
                MakeServiceOversikt(true, bwAutoRanking);
            }

            string file = settingsPath + "\\rankingQuick.html";
            if (!File.Exists(file))
            {
                Logg.Log("Kveldstall: Fant ikke kveldsranking fil. Avbryter..", Color.Red);
                e.Result = 6;
                return;
            }

            Logg.Log("Kveldstall: Konverterer kveldranking til PDF..");
            string filePdf = CreatePDF("Quick", "", bwPDF);

            KgsaEmail email = new KgsaEmail(this);
            email.Send(filePdf, DateTime.Now, "Quick", appConfig.epostEmneQuick, appConfig.epostBodyQuick);

            if (filePdf == "")
            {
                // Feil under pdf generering
                Logg.Log("Kveldstall: Feil oppstod under konvertering og sending av kveldsranking til PDF. Se logg for detaljer.", Color.Red);
                e.Result = 6;
                return;
            }
        }

        private void bwQuickAuto_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressStop();
            autoMode = false;
            if (e.Error != null)
            {
                Logg.Log("Kveldstall: Ukjent feil oppstod. " + e.Error.Message + " Se logg for detaljer.", Color.Red);
            }
            else if (e.Cancelled || (int)e.Result == 2)
            {
                Logg.Log("Kveldstall: Avbrutt av bruker.");
            }
            else if ((int)e.Result != 0)
            {
                Logg.Log("Kveldstall: Ukjent feil oppstod. Se logg for detaljer.", Color.Red);
            }
            else
            {
                Logg.Log("Kveldstall: Fullført uten feil.", Color.Green);
            }
            processing.HideDelayed();
            this.Activate();
        }

        private List<string> MakeBudgetHtml(DataTable table)
        {
            try
            {
                var doc = new List<string>();

                Logg.Debug("Henter dagens budsjett fra database..");
                DataTable tableBudget = database.tableDailyBudget.GetBudgetFromDate(appConfig.Avdeling, DateTime.Now);
                if (tableBudget == null || tableBudget.Rows.Count < 6 || table == null || table.Rows.Count < 6)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Mangler data for budsjett</span><br>");
                    return doc;
                }

                doc.Add("<br><table style='width:100%'><tr><td>");
                doc.Add("<span class='Subtitle'>Resultat mot budsjett " + avdeling.Get(appConfig.Avdeling)
                    + " - " + DateTime.Now.ToString("dddd d. MMMM yyyy  HH:mm", norway) + "</span>");

                doc.Add("<table class='OutertableNormal'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.Add("<thead><tr>");

                doc.Add("<th class=\"{sorter: 'text'}\" style='background:#bfd2e2;' width=95 >Category</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;border-left:2px solid #000;' width=60 >Sales</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;' width=60 >Budget</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background:#bfd2e2;' width=60 >Index</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;border-left:2px solid #000;' width=60 >GM %</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;' width=60 >Budget</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background:#bfd2e2;' width=60 >Variance</td>");

                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;border-left:2px solid #000;' width=60 >GM</td>");
                doc.Add("<th class=\"{sorter: 'digit'}\" style='background:#bfd2e2;' width=60 >Budget</td>");
                doc.Add("<th class=\"{sorter: 'procent'}\" style='background:#bfd2e2;' width=60 >Variance</td>");

                doc.Add("</tr></thead>");
                doc.Add("<tbody>");

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (Convert.ToInt32(table.Rows[i]["Favoritt"]) != appConfig.Avdeling)
                        break;

                    string kategori = table.Rows[i]["Avdeling"].ToString();
                    decimal omset_budget = 0, inntjen_budget = 0, margin_budget = 0;
                    decimal omset = Convert.ToDecimal(table.Rows[i]["Omsetn"]);
                    decimal inntjen = Convert.ToDecimal(table.Rows[i]["Fortjeneste"]);
                    decimal margin = Convert.ToDecimal(table.Rows[i]["Margin"]);
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
                        doc.Add("</tbody><tfoot><tr><td class='text-cat'>" + kategori + "</td>");
                    else
                        doc.Add("<tr><td class='text-cat'>" + kategori + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;text-align: right;' >" + tools.NumberStyle_Normal(omset) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Normal(omset_budget) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Percent(omset, omset_budget, true, true) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;text-align: right;' >" + tools.NumberStyle_Normal(margin, 1, "%") + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Normal(margin_budget, 1, "%") + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Normal(margin - margin_budget, 1, "%", true, true) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;text-align: right;' >" + tools.NumberStyle_Normal(inntjen) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Normal(inntjen_budget) + "</td>");
                    doc.Add("<td class='numbers-gen'>" + tools.NumberStyle_Normal(inntjen - inntjen_budget, 0, "", true, true) + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</table></td></tr></table>");

                doc.Add("</td></tr></table>");

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Uventet feil oppstod under generering av daglig budsjett tabellen", Color.Red);
            }
            return new List<string> { };
        }

        public bool MakeQuickHtml(DataTable table)
        {
            try
            {
                var doc = new List<string>();
                GetHtmlStart(doc, true);

                if (table == null || table.Rows.Count < 6)
                {
                    Logg.Log("Mangler data for generering av kveldstall", Color.Red);
                    return false;
                }

                doc.Add("<span class='Title'>Kveldstall (" + avdeling.Get(appConfig.Avdeling) + ")</span>");
                doc.Add("<span class='Generated'>Ranking generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span><br>");

                for (int f = 0; f < table.Rows.Count; f++)
                {
                    if (appConfig.dailyBudgetIncludeInQuickRanking && f == 0)
                        doc.AddRange(MakeBudgetHtml(table));

                    int intAvd = Convert.ToInt32(table.Rows[f][0]);
                    doc.Add("<br><table style='width:100%'><tr><td>");
                    doc.Add("<span class='Subtitle'>Dagens tall for " + avdeling.Get(Convert.ToInt32(table.Rows[f]["Favoritt"].ToString()))
                        + " - " + DateTime.Now.ToString("dddd d. MMMM yyyy  HH:mm", norway) + " </span><br>");
                    doc.Add("<table class='OutertableNormal'><tr><td>");
                    doc.Add("<table class='tablesorter'>");
                    doc.Add("<thead><tr>");
                    doc.Add("<th class=\"{sorter: 'text'}\" width=95 >Category</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=55 >Salg</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Omset.</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Fritt</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Inntjen.</td>");
                    doc.Add("<th class=\"{sorter: 'procent'}\" width=60 >Margin</td>");
                    doc.Add("<th class=\"{sorter: 'digit'}\" width=60 >Rabatt</td>");

                    if (appConfig.macroImportQuickSales && intAvd == appConfig.Avdeling && this.kvdelstallInfo != null)
                    {
                        if (this.kvdelstallInfo.Salg.Count > 0)
                        {
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 style='background:#80c34a;'>RTG/SA</td>");
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=65 style='background:#80c34a;'>Omset.</td>");
                            doc.Add("<th class=\"{sorter: 'digit'}\" width=65 style='background:#80c34a;'>Inntjen.</td>");
                            doc.Add("<th class=\"{sorter: 'procent'}\" width=55 style='background:#80c34a;'>" +
                                "<abbr title='Btokr. inntjen. Tjenester / Btokr. inntjen. alle varer'>%</abbr></td>");
                        }
                    }

                    doc.Add("</tr></thead>");
                    doc.Add("<tbody>");

                    double KgsaAntallTot = 0, KgsaBtokrTot = 0, KgsaSalgsprisTot = 0;
                    for (int i = f; i < table.Rows.Count; i++)
                    {
                        if (intAvd != (Convert.ToInt32(table.Rows[i][0])))
                        {
                            f--;
                            break;
                        }

                        f++;
                        if (table.Rows.Count == i + 1 || table.Rows[i]["Avdeling"].ToString() == "TOTALT") // siste row
                            doc.Add("</tbody><tfoot><tr><td class='text-cat'>" + table.Rows[i]["Avdeling"] + "</td>");
                        else
                            doc.Add("<tr><td class='text-cat'>" + table.Rows[i]["Avdeling"] + "</td>");
                        doc.Add("<td class='numbers-small'>" + PlusMinus(table.Rows[i]["Salg"]) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Omsetn"]) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Fritt"]) + "</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Fortjeneste"]) + "</td>");
                        doc.Add("<td class='numbers-percent'>" + table.Rows[i]["Margin"] + " %</td>");
                        doc.Add("<td class='numbers-gen'>" + PlusMinus(table.Rows[i]["Rabatt"]) + "</td>");

                        if (appConfig.macroImportQuickSales && intAvd == appConfig.Avdeling && this.kvdelstallInfo != null)
                        {
                            if (this.kvdelstallInfo.Salg.Count > 0)
                            {
                                if ("TOTALT" == table.Rows[i]["Avdeling"].ToString())
                                {
                                    doc.Add("<td class='numbers-service'>" + PlusMinus(KgsaAntallTot) + "</td>");
                                    doc.Add("<td class='numbers-gen'>" + PlusMinus(KgsaSalgsprisTot) + "</td>");
                                    doc.Add("<td class='numbers-gen'>" + PlusMinus(KgsaBtokrTot) + "</td>");
                                    doc.Add("<td class='numbers-percent'>" + Math.Round(KgsaBtokrTot / Convert.ToDouble(table.Rows[i]["Fortjeneste"].ToString()) * 100, 2) + " %</td>");
                                }
                                else
                                {
                                    bool found = false;
                                    foreach (KveldstallArray item in kvdelstallInfo.Salg)
                                    {
                                        if (item.Type == table.Rows[i]["Avdeling"].ToString())
                                        {
                                            doc.Add("<td class='numbers-service'>" + PlusMinus(item.Antall) + "</td>");
                                            KgsaAntallTot += item.Antall;
                                            doc.Add("<td class='numbers-gen'>" + PlusMinus(item.Salgspris) + "</td>");
                                            KgsaSalgsprisTot += item.Salgspris;
                                            doc.Add("<td class='numbers-gen'>" + PlusMinus(item.Btokr) + "</td>");
                                            KgsaBtokrTot += item.Btokr;
                                            doc.Add("<td class='numbers-percent'>" + Math.Round(item.Btokr / Convert.ToDouble(table.Rows[i]["Fortjeneste"].ToString()) * 100, 2) + " %</td>");
                                            found = true;
                                            break;
                                        }
                                    }
                                    if (!found)
                                    {
                                        doc.Add("<td class='numbers-service'>" + PlusMinus(0) + "</td>");
                                        doc.Add("<td class='numbers-gen'>" + PlusMinus(0) + "</td>");
                                        doc.Add("<td class='numbers-gen'>" + PlusMinus(0) + "</td>");
                                        doc.Add("<td class='numbers-percent'>-</td>");
                                    }
                                }
                            }
                        }

                        doc.Add("</tr>");
                    }

                    doc.Add("</tfoot></table></td></tr></table>");
                    doc.Add("</td></tr></table>");
                }
                doc.Add(Resources.htmlEnd);
                File.WriteAllLines(htmlRankingQuick, doc.ToArray(), Encoding.Unicode);

                webHTML.Navigate(htmlRankingQuick);

                return true;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                Logg.Log("Uventet feil under generering av side for kveldstall. Feilmelding: " + ex.Message, Color.Red);
            }
            return false;
        }


        public string PlusMinus(object arg)
        {
            try
            {
                decimal var = Math.Round(Convert.ToDecimal(arg), 0);
                string value = appConfig.visningNull;
                if (var < 0)
                    value = "<span style='color:red'>" + var.ToString("#,##0") + "</span>";
                if (var > 0)
                    value = var.ToString("#,##0");
                return value;
            }
            catch
            {
                return appConfig.visningNull;
            }
        }
    }
}
