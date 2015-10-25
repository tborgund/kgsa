using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace KGSA
{
    public class OpenXML
    {
        FormMain main;
        private List<OpenXMLDocument> documentDatabase;

        public OpenXML(FormMain form)
        {
            this.main = form;
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
                FormError errorMsg = new FormError("Feil ved åpning av regneark", ex);
                errorMsg.ShowDialog();
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
