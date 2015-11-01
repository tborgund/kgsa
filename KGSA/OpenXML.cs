using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;

namespace KGSA
{
    public class OpenXML
    {
        FormMain main;
        private List<OpenXMLDocument> documentDatabase;
        public static string DOC_ALL_SALES_REP = "AllSalesRep";

        public OpenXML(FormMain form)
        {
            main = form;
            documentDatabase = new List<OpenXMLDocument> { };
            LoadDatabase();
        }

        public void ClearDatabase()
        {
            documentDatabase = null;
            documentDatabase = new List<OpenXMLDocument> { };
            try
            {
                File.Delete(FormMain.settingsPath + @"\OpenXMLdb.dat");
            }
            catch { }
        }

        public void SaveDatabase()
        {
            try
            {
                if (documentDatabase != null)
                    if (documentDatabase.Count > 0)
                        using (Stream stream = File.Open(FormMain.settingsPath + @"\OpenXMLdb.dat", FileMode.Create))
                        {
                            BinaryFormatter bin = new BinaryFormatter();
                            bin.Serialize(stream, documentDatabase);
                            Log.d("OpenXML: Fullført lagring av databasen.");
                        }
            }
            catch (IOException ex)
            {
                Log.d("OpenXML: Kritisk feil ved lagring av databasen.", ex);
            }
        }

        public void LoadDatabase()
        {
            try
            {
                if (File.Exists(FormMain.settingsPath + @"\OpenXMLdb.dat"))
                {
                    using (Stream stream = File.Open(FormMain.settingsPath + @"\OpenXMLdb.dat", FileMode.Open))
                    {
                        BinaryFormatter bin = new BinaryFormatter();
                        documentDatabase = (List<OpenXMLDocument>)bin.Deserialize(stream);
                        Log.d("OpenXML: Fullført åpning av databasen.");
                    }
                }
            }
            catch (IOException ex)
            {
                Log.d("OpenXML: Kritisk feil ved åpning av databasen.", ex);
                try
                {
                    File.Delete(FormMain.settingsPath + @"\OpenXMLdb.dat");
                }
                catch { }
            }
        }

        public void OpenDocument(DateTime date)
        {
            try
            {
                string tab = main.readCurrentTab();
                string page = main.currentPage();

                if ((tab.Equals("Ranking") || tab.Equals("Avdelinger") || tab.Equals("Store") || tab.Equals("Budget")) && !String.IsNullOrEmpty(page))
                {
                    string file = ExportDocument(page, date);
                    if (file != null)
                        System.Diagnostics.Process.Start(file);
                    else
                        Log.n("Regneark for gjeldene side/dato mangler, eller finnes ikke. Oppdater og forsøk igjen.", Color.Red);
                    return;
                }
                else
                    Log.n("Kan ikke åpne regneark for gjeldene side. Velg en annen ranking eller oppdater.", Color.Red);
            }
            catch(Exception ex)
            {
                Log.ErrorDialog(ex, "Feil ved åpning av regneark", "KGSA OpenXML");
            }
        }

        public void SaveDocument(DataTable table, string bookName, string sheetName, DateTime date, string header = "", string[] ignoreColumns = null)
        {
            if (table == null || bookName == null || sheetName == null)
            {
                Log.d("OpenXML: Mangler argumenter for lagring av dokument.");
                return;
            }

            if (table.Rows.Count == 0)
            {
                Log.d("OpenXML: Tabellen har ingen linjer, kan ikke opprette side.");
                return;
            }

            if (String.IsNullOrEmpty(bookName) || String.IsNullOrEmpty(sheetName))
            {
                Log.d("OpenXML: Navn ikke angitt for dokumentet og/eller siden.");
                return;
            }

            bool foundBook = false, foundSheet = false;
            for (int i = 0; i < documentDatabase.Count; i++)
            {
                //Logg.Debug("OpenXML: Searching documentDatabase.. :" + documentDatabase[i].date.ToShortDateString() + " - " + documentDatabase[i].dataset.DataSetName);
                if (documentDatabase[i].dataset.DataSetName.Equals(bookName) && documentDatabase[i].date.Date == date.Date)
                {
                    foundBook = true;
                    for (int d = 0; d < documentDatabase[i].dataset.Tables.Count; d++)
                    {
                        //Logg.Debug("OpenXML: Searching documentDatabase.dataset.Tables.. [" + documentDatabase[i].dataset.DataSetName + "]: " + documentDatabase[i].dataset.Tables[d].TableName);
                        if (documentDatabase[i].dataset.Tables[d].TableName.Equals(sheetName))
                        {
                            foundSheet = true;
                            documentDatabase[i].dataset.Tables.RemoveAt(d);
                        }
                    }

                    var sheet = TrimColumns(table.Copy(), ignoreColumns);
                    sheet.TableName = sheetName;
                    sheet.Namespace = header;
                    documentDatabase[i].dataset.Tables.Add(sheet);

                    if (foundSheet)
                        Log.d("OpenXML: Overskrev ark '" + sheetName + "' tilhørende dokument '" + bookName + "' - Ark totalt: " + documentDatabase[i].dataset.Tables.Count);
                    else
                        Log.d("OpenXML: Lagt til nytt ark med navn '" + sheetName + "' til dokument '" + bookName + "' - Ark totalt: " + documentDatabase[i].dataset.Tables.Count);
                    return;
                }
            }

            if (!foundBook)
            {
                var document = new OpenXMLDocument();
                document.dataset = new DataSet();
                document.dataset.DataSetName = bookName;
                document.date = date;

                var sheet = TrimColumns(table.Copy(), ignoreColumns);
                sheet.TableName = sheetName;
                sheet.Namespace = header;
                document.dataset.Tables.Add(sheet);
                documentDatabase.Add(document);

                Log.d("OpenXML: Opprettet nytt dokument med navn '" + bookName + "' og første ark med navn '" + sheetName + " - Linjer: " + table.Rows.Count);
            }
        }

        public void DeleteDocument(string bookName, DateTime date)
        {
            if (documentDatabase.Count == 0)
                return;
            for (int i = 0; i < documentDatabase.Count; i++)
            {
                if (documentDatabase[i].dataset.DataSetName.Equals(bookName) && documentDatabase[i].date.Date == date.Date)
                {
                    documentDatabase.RemoveAt(i);
                    Log.d("OpenXML: Fjernet dokument med navn '" + bookName + "' " + date.ToShortDateString());
                    return;
                }
            }
            Log.d("OpenXML: Dokument med navn '" + bookName + "' fantes ikke fra dør.");
        }

        private string ExportDocument(string bookName, DateTime date)
        {
            for (int i = 0; i < documentDatabase.Count; i++)
            {
                if (documentDatabase[i].dataset.DataSetName.Equals(bookName) && documentDatabase[i].date.Date == date.Date)
                    return CreateDocument(documentDatabase[i].dataset, bookName, date);
            }
            Log.d("Fant ikke dokumentet '" + bookName + "' for lagring!");
            return null;
        }

        private string CreateDocument(DataSet dataset, string bookName, DateTime date)
        {
            try
            {
                Log.d("OpenXML: Genererer dokument '" + bookName + "'...");

                var wb = new XLWorkbook();

                for (int i = 0; i < dataset.Tables.Count; i++)
                {
                    var ws = wb.Worksheets.Add(dataset.Tables[i].TableName);

                    ws.Cell(1, 1).Value = "Eksportert av KGSA " + FormMain.version + " - " + DateTime.Now.ToShortDateString();
                    ws.Range(1, 1, 1, 4).Merge().AddToNamed("Kgsa");

                    ws.Cell(2, 1).Value = dataset.Tables[i].Namespace;
                    ws.Range(2, 1, 2, 4).Merge().AddToNamed("Headers");

                    var tableWithData = ws.Cell(3, 1).InsertTable(dataset.Tables[i].AsEnumerable());

                    ws.Columns().AdjustToContents();
                }

                // Prepare the style for the titles
                var kgsaStyle = wb.Style;
                kgsaStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                var headersStyle = wb.Style;
                headersStyle.Font.Bold = true;
                headersStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                // Format all titles in one shot
                wb.NamedRanges.NamedRange("Kgsa").Ranges.Style = kgsaStyle;
                wb.NamedRanges.NamedRange("Headers").Ranges.Style = headersStyle;

                string filename = FormMain.settingsTemp + @"\KGSA_" + bookName + "_" + date.ToShortDateString() + ".xlsx";

                wb.SaveAs(filename);
                Log.n("OpenXML: Generert dokument '" + bookName + "' (" + dataset.Tables.Count + " sider) file://" + filename.Replace(' ', (char)160), Color.Green);
                return filename;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
            }
            return null;
        }

        /// <summary>
        /// Creates and opens special OpenXML documents, runs async
        /// </summary>
        public void CreateAndOpenXml(string name)
        {
            main.worker = new BackgroundWorker();
            main.worker.DoWork += new DoWorkEventHandler(openXML_DoWork);
            main.worker.ProgressChanged += new ProgressChangedEventHandler(main.bwProgressCustom_ProgressChanged);
            main.worker.WorkerReportsProgress = true;
            main.worker.WorkerSupportsCancellation = true;
            main.worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(openXML_Completed);
            main.processing.SetVisible = true;
            main.processing.SetProgressStyle = ProgressBarStyle.Marquee;
            main.processing.SetBackgroundWorker = main.worker;
            main.worker.RunWorkerAsync(name);
        }

        private void openXML_DoWork(object sender, DoWorkEventArgs e)
        {
            main.ProgressStart();
            e.Result = CreateDocumentSpecial((string)e.Argument, main.worker);

            if (main.worker.CancellationPending)
                e.Cancel = true;
        }

        private void openXML_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            main.ProgressStop();
            main.processing.SetProgressStyle = ProgressBarStyle.Continuous;
            main.processing.SetValue = 100;

            if (!e.Cancelled && e.Error == null && e.Result != null)
            {
                string path = (string)e.Result;
                try
                {
                    Log.n("Åpner regneark..: " + Path.GetFileName(path));
                    Process.Start(path);

                    main.processing.SetText = "Ferdig!";
                }
                catch (FileNotFoundException)
                {
                    Log.Alert("Dokumentet " + path + " finnes ikke.\nFilen kan være slettet, låst eller ødelagt");
                    main.processing.SetText = "Avbrutt";
                }
                catch (Exception ex)
                {
                    main.processing.SetText = "Avbrutt!";
                    Log.Unhandled(ex);
                    Log.e("Klarte ikke åpne dokument. Feilmelding: " + ex.Message);
                }
                main.processing.HideDelayed();
            }
            else if (e.Cancelled)
            {
                main.processing.SetText = "Avbrutt!";
                main.processing.HideDelayed();
                Log.e("Avbrutt av bruker");
            }
            else
            {
                main.processing.Visible = false;
            }
        }

        public string CreateDocumentSpecial(string selection, BackgroundWorker bw)
        {
            try
            {
                if (selection.IsNullOrWhiteSpace() || bw == null)
                {
                    Log.e("OpenXML: Mangler argumenter for å lage dokument");
                    return "";
                }

                if (selection.Equals(DOC_ALL_SALES_REP))
                    return CreateDocument_AllSalesRep(bw);
                else
                    Log.e("OpenXML: Ukjent navn på dokument '" + selection + "'");
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("OpenXML: Uventet feil - klarte ikke lage dokument " + selection + ". Feilmelding: " + ex.Message);
            }
            return "";
        }

        private string CreateDocument_AllSalesRep(BackgroundWorker bw)
        {
            main.processing.SetText = "Lager regneark: Forbereder..";

            PageBudgetAllSales page = new PageBudgetAllSales(main, true, bw, main.webBudget);
            DateTime dateStart = main.appConfig.dbTo.EndOfLastWorkWeek(main.appConfig.ignoreSunday);
            XLWorkbook workBook = new XLWorkbook();
            try
            {
                for (int i = 0; i < 30; i = i + 7)
                {
                    DateTime date = dateStart.AddDays(-i);
                    if (date < main.appConfig.dbFrom)
                    {
                        Log.d("OpenXML: Skipping sheet, don't have complete data for week " + main.database.GetIso8601WeekOfYear(date)
                            + ". Database goes from " + main.appConfig.dbFrom.ToShortDateString() + " to " + main.appConfig.dbTo.ToShortDateString());
                        break;
                    }

                    main.processing.SetText = "Lager regneark: Genererer ark for uke " + main.database.GetIso8601WeekOfYear(date) + "..";
                    Log.n("OpenXML: Lager side for uke " + main.database.GetIso8601WeekOfYear(date) + "..");

                    IXLWorksheet ws;
                    if (i == 0)
                        ws = workBook.Worksheets.Add("Uke " + main.database.GetIso8601WeekOfYear(date).ToString()).SetTabColor(XLColor.Green);
                    else
                        ws = workBook.Worksheets.Add("Uke " + main.database.GetIso8601WeekOfYear(date).ToString());

                    DataTable table = page.MakeTableForWeek(main.appConfig.Avdeling, date);
                    int width = table.Columns.Count;
                    int height = table.Rows.Count;

                    // ws.Range(Row, Colum, LastRow, LastColum)
                    var tableWithData = ws.Cell(3, 1).InsertTable(table.AsEnumerable());

                    ws.Cell(1, 1).Value = "Eksportert av KGSA " + FormMain.version + " - " + DateTime.Now.ToString("dddd d. MMMM  HH:mm", FormMain.norway);
                    ws.Range(1, 1, 1, width).Merge().AddToNamed("HeaderLeft");

                    ws.Cell(2, 1).Value = "Selgeroversikt - Uke " + main.database.GetIso8601WeekOfYear(date);
                    ws.Range(2, 1, 2, 2).Merge().AddToNamed("HeaderLeft");

                    ws.Cell(2, 3).Value = date.StartOfWeek().ToString("dddd d. MMMM", FormMain.norway) + " - " + date.ToString("dddd d. MMMM", FormMain.norway);
                    ws.Range(2, 3, 2, 8).Merge().AddToNamed("HeaderRight");

                    ws.Range(4, 3, height + 3, 3).AddToNamed("Numbers");
                    ws.Range(4, 4, height + 3, 4).AddToNamed("Percent");
                    ws.Range(4, 5, height + 3, 5).AddToNamed("Numbers");
                    ws.Range(4, 6, height + 3, 6).AddToNamed("InputHours");
                    ws.Range(4, 7, height + 3, 7).AddToNamed("Numbers");
                    ws.Range(4, 8, height + 3, 8).AddToNamed("Numbers");

                    for (int d = 0; d < height; d++)
                    {
                        ws.Cell(4 + d, 7).FormulaA1 = "=IF(F" + (4 + d) + "=0,\"0\",E" + (4 + d) + "/F" + (4 + d) + ")";
                        ws.Cell(4 + d, 8).FormulaA1 = "=IF(F" + (4 + d) + "=0,\"0\",C" + (4 + d) + "/F" + (4 + d) + ")";
                    }

                    ws.Columns().AdjustToContents();
                    ws.Columns("9:34").Hide();
                    ws.Cell(4, 6).Select();
                }
                
                for (int i = 0; i < 4; i++)
                {
                    DateTime date = main.appConfig.dbTo.AddMonths(-i);
                    var firstDay = new DateTime(date.Year, date.Month, 1);
                    var lastDay = firstDay.AddMonths(1).AddDays(-1);
                    if (firstDay < main.appConfig.dbFrom)
                    {
                        Log.d("OpenXML: Skipping sheet, don't have complete data for " + date.ToString("MMMM yyyy")
                            + ". Database goes from " + main.appConfig.dbFrom.ToShortDateString() + " to " + main.appConfig.dbTo.ToShortDateString());
                        break;
                    }

                    main.processing.SetText = "Lager regneark: Genererer ark for måned " + date.ToString("MMMM yyyy") + "..";
                    Log.n("OpenXML: Lager side for måned " + date.ToString("MMMM yyyy") + "..");

                    IXLWorksheet ws;
                    string sheetName = date.ToString("MMMM");
                    if (date.Month == main.appConfig.dbTo.Month && date.Year == main.appConfig.dbTo.Year)
                    {
                        lastDay = main.appConfig.dbTo;
                        sheetName = "MTD " + sheetName;
                        ws = workBook.Worksheets.Add(sheetName).SetTabColor(XLColor.Green);
                    }
                    else
                        ws = workBook.Worksheets.Add(sheetName);

                    DataTable table = page.MakeTableForMonth(main.appConfig.Avdeling, date);
                    int width = table.Columns.Count;
                    int height = table.Rows.Count;

                    // ws.Range(Row, Colum, LastRow, LastColum)
                    var tableWithData = ws.Cell(3, 1).InsertTable(table.AsEnumerable());

                    ws.Cell(1, 1).Value = "Eksportert av KGSA " + FormMain.version + " - " + DateTime.Now.ToString("dddd d. MMMM  HH:mm", FormMain.norway);
                    ws.Range(1, 1, 1, width).Merge().AddToNamed("HeaderLeft");

                    ws.Cell(2, 1).Value = "Selgeroversikt - Måned " + date.ToString("MMMM yyyy", FormMain.norway);
                    ws.Range(2, 1, 2, 2).Merge().AddToNamed("HeaderLeft");

                    ws.Cell(2, 3).Value = firstDay.ToString("dddd d. MMMM", FormMain.norway) + " - " + lastDay.ToString("dddd d. MMMM", FormMain.norway);
                    ws.Range(2, 3, 2, 8).Merge().AddToNamed("HeaderRight");

                    ws.Range(4, 3, height + 3, 3).AddToNamed("Numbers");
                    ws.Range(4, 4, height + 3, 4).AddToNamed("Percent");
                    ws.Range(4, 5, height + 3, 5).AddToNamed("Numbers");
                    ws.Range(4, 6, height + 3, 6).AddToNamed("InputHours");
                    ws.Range(4, 7, height + 3, 7).AddToNamed("Numbers");
                    ws.Range(4, 8, height + 3, 8).AddToNamed("Numbers");

                    for (int d = 0; d < height; d++)
                    {
                        ws.Cell(4 + d, 7).FormulaA1 = "=IF(F" + (4 + d) + "=0,\"0\",E" + (4 + d) + "/F" + (4 + d) + ")";
                        ws.Cell(4 + d, 8).FormulaA1 = "=IF(F" + (4 + d) + "=0,\"0\",C" + (4 + d) + "/F" + (4 + d) + ")";
                    }

                    ws.Columns().AdjustToContents();
                    ws.Columns("9:34").Hide();
                    ws.Cell(4, 6).Select();
                }

                main.processing.SetText = "Lager regneark: Snart der!";
                Log.n("OpenXML: Lager regneark..");

                try
                {
                    IXLStyle sHeaderLeft = workBook.Style;
                    sHeaderLeft.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    workBook.NamedRanges.NamedRange("HeaderLeft").Ranges.Style = sHeaderLeft;

                    IXLStyle sHeaderRight = workBook.Style;
                    sHeaderRight.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    workBook.NamedRanges.NamedRange("HeaderRight").Ranges.Style = sHeaderRight;

                    IXLStyle sNumbers = workBook.Style;
                    sNumbers.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    sNumbers.NumberFormat.Format = "#,##0";
                    workBook.NamedRanges.NamedRange("Numbers").Ranges.Style = sNumbers;

                    IXLStyle sPercent = workBook.Style;
                    sPercent.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    sPercent.NumberFormat.Format = "0.0%";
                    workBook.NamedRanges.NamedRange("Percent").Ranges.Style = sPercent;

                    IXLStyle sHourInput = workBook.Style;
                    sHourInput.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sHourInput.Fill.BackgroundColor = XLColor.FromHtml("#ffcc99");
                    sHourInput.Font.FontColor = XLColor.FromHtml("#3f3f76");
                    sHourInput.NumberFormat.Format = "#,##0";
                    workBook.NamedRanges.NamedRange("InputHours").Ranges.Style = sHourInput;
                }
                catch (Exception ex)
                {
                    Log.Unhandled(ex);
                    Log.n("OpenXML: Exception while attempting to apply styles to document. It might not be fatal", Color.Brown);
                }

                string path = main.tools.GetTempFilename("KGSA Selgeroversikt " + main.appConfig.dbTo.ToShortDateString(), ".xlsx");

                workBook.SaveAs(path);
                Log.n("OpenXML: Opprettet dokument: file://" + path.Replace(' ', (char)160), Color.Green);

                return path;
            }
            catch (Exception ex)
            {
                Log.Unhandled(ex);
                Log.e("OpenXML: Uventet feil - klarte ikke lage dokumentet. Feilmelding: " + ex.Message);
            }
            finally
            {
                if (workBook != null)
                    workBook = null;
                if (page != null)
                    page = null;
            }
            return "";
        }

        private DataTable TrimColumns(DataTable table, string[] ignoreColumns)
        {
            if (ignoreColumns != null)
            {
                List<string> ls = new List<string> { };
                foreach (string col in ignoreColumns)
                    foreach (DataColumn dc in table.Columns)
                        if (col.Equals(dc.ColumnName))
                            ls.Add(dc.ColumnName);

                foreach (string str in ls)
                    table.Columns.Remove(str);

            }
            return table;
        }
    }

    [Serializable]
    public class OpenXMLDocument
    {
        public DataSet dataset = new DataSet();
        public DateTime date = FormMain.rangeMin;
    }
}
