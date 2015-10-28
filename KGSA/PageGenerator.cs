using KGSA.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;

namespace KGSA
{
    public class PageGenerator
    {
        public FormMain main;
        public bool runningInBackground = false;
        public BackgroundWorker worker;
        public List<string> doc;
        public System.Windows.Forms.WebBrowser browser;
        public DateTime pickedDate = FormMain.rangeMin;

        public static string Class_Style_Generic = "numbers-gen";
        public static string Class_Style_Small = "numbers-small";
        public static string Class_Style_Percent = "numbers-percent";
        public static string Class_Style_Text_Cat = "text-cat";

        public static string Sorter_Type_Text = "text";
        public static string Sorter_Type_Digit = "digit";
        public static string Sorter_Type_Procent = "procent";

        public static string Color_Bg_CoolBlue = "background:#bfd2e2;";

        protected void ShowProgress()
        {
            doc.Add("<span class='Loading'>Beregner..</span>");
            if (!runningInBackground && main.timewatch.ReadyForRefresh())
                browser.DocumentText = string.Join(null, doc.ToArray());
            doc.RemoveAt(doc.Count - 1);
        }

        protected void AddPage_Title(string title)
        {
            doc.Add("<h1>" + title + "</h1>");
            doc.Add("<span class='Generated'>Side generert: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "</span>");
            doc.Add("<br>");
        }

        protected void AddPage_SubTitle(string title)
        {
            doc.Add("<h2>" + title + "</h2>");
        }

        protected void AddPage_Header(string title)
        {
            doc.Add("<h3>" + title + "</h3>");
        }

        protected void AddTable_Start(string title = "", string className = "OutertableNormal")
        {
            doc.Add("<div class='no-break'>");
            //doc.Add("<br><table style='width:100%'><tr><td>");
            AddPage_Header(title);
            doc.Add("<table class='" + className + "'><tr><td>");
            doc.Add("<table class='tablesorter'>");
        }

        protected void AddTable_End()
        {
            doc.Add("</table>");
            doc.Add("</td></tr></table>");
            doc.Add("</div>");
        }

        protected void AddTable_Header_Start()
        {
            doc.Add("<thead><tr>");
        }

        protected void AddTable_Header_Name(string name, int width, string style = "", string sortType = "text", string hoverText = "")
        {
            if (!String.IsNullOrEmpty(hoverText))
                name = "<abbr title='" + hoverText + "'>" + name + "</abbr>";
            if (!String.IsNullOrEmpty(style))
                style = " style='" + style + "' ";
            doc.Add("<th class='{sorter: '" + sortType + "' width=" + width + " " + style + " >" + name + "</td>");
        }

        protected void AddTable_Header_End()
        {
            doc.Add("</tr></thead>");
        }

        protected void AddTable_Body_Start()
        {
            doc.Add("<tbody>");
        }

        protected void AddTable_Row_Start()
        {
            doc.Add("<tr>");
        }

        protected void AddTable_Footer_Start()
        {
            doc.Add("<tfoot>");
        }

        protected void AddTable_Footer_End()
        {
            doc.Add("</tfoot>");
        }

        protected void AddTable_Row_Cell(string content, string style = "", string className = "numbers-gen")
        {
            if (!String.IsNullOrEmpty(style))
                style = " style='" + style + "' ";
            doc.Add("<td class='" + className + "' " + style + ">" + content + "</td>");
        }

        protected void AddTable_Row_End()
        {
            doc.Add("</tr>");
        }

        protected void AddTable_Body_End()
        {
            doc.Add("</tbody>");
        }

        protected void AddPage_Start(bool watermark = false, string title = "Untitled")
        {
            doc.Add("<html>");
            doc.Add("<head>");
            doc.Add("<meta charset=\"UTF-8\">");
            doc.Add("<title>KGSA - " + title + "</title>");

            doc.Add("<style id=\"stylesheet\">");
            doc.Add("body {");
            doc.Add("    font-weight:400;");
            doc.Add("    font-family:Calibri, sans-serif;");
            doc.Add("    font-style:normal;");
            doc.Add("    text-decoration:none;");
            doc.Add("    display: block;");
            doc.Add("    padding: 0;");
            doc.Add("    margin: 10px;");
            if (watermark)
            {
                doc.Add("    background-repeat:no-repeat;");
                doc.Add("    background-attachment:fixed;");
                doc.Add("    background-position: top right;");
                if (main.appConfig.chainElkjop)
                    doc.Add("    background-image:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUAAAAIRCAMAAAAMdJRXAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6MkVCMEIxQTI5MTIzMTFFMzlGRDVCRERCMDgyOUMyRTIiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6MkVCMEIxQTM5MTIzMTFFMzlGRDVCRERCMDgyOUMyRTIiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDoyRUIwQjFBMDkxMjMxMUUzOUZENUJEREIwODI5QzJFMiIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDoyRUIwQjFBMTkxMjMxMUUzOUZENUJEREIwODI5QzJFMiIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PriKAVcAAAMAUExURdbow/3+/d3swenz7dzsyMnj1eTx6sHgzvn79d3syuPw4trr29npzLXWxr/czt7txODt2dbozPr8+Nrq09npwcLexNnpxe318ebx0fT57eDtxfb68b3dytTo3LbZxdHlwbjYyPD26ubx1u724uHv5unz2NHlyfX59fz9+ujy3NnqyvD25cbh0sziwrjcxsrhwujy1sjgxdTmwfL49dzrvtrrvtHm28Dgzubw2uz03r3dxejy4r3bzMXgxdHlxOTwzdnq4s3k2Ony5t7t5sTf0tTmxbjZxvb6+Njpv9bov+HuzcnhyPL46dTmysTgycHgys/kwrzay+715dvqvc3jw97sytbo08Lez9rqx97u4ev03dvrxb7dw9nqv77czLrcyLrayvP47czjysHfx87kxdDm0uPvy+/28dDkztfox9box+rz2vH48trs4b7ey7rayNzrwcvkzPj7+brZyrjbxMrjzrbYyNzqycDezdfozs7l0c7kyvr8+8jizrrYyeDuy9Xnv7/czczjzvH35sbhzNbnzdPmz83kz7nayN/tzdXoyqvRv9/tyeDuyNvswcHgzbvcxNzrytrqy8HezsHfztXnzN7sydPmzNfoy9jpy9npy8Thzd7tyc7kzdTnzMDdzszkzdLmzNHlzdDlzeDuyc/kzcvjzcrjzdvqysnjzdvrysfizcbizcDczsPhzdLlzNTmzNLlzdvqy8jizcHdztboy9fpy8XhzcLhzcXizc3kzcDdz+Huyd7tyNzry93ry8Dcz9LlztPmzcHdz+DtytXnzbbXx/T569npysHez9/tyNTmzev04dzqy8fj0MHcz97sx8vjzNvrzOLuydTnzdzqzcDeztHmzsHcztboysjjzMniztPlzcrjzNrpycPgztbp4LnZydvsxtjpydXoyMXhzsrgw9fpyc/ky9PnzcjjzcnizMnjzNfptdzswNjpt9Hky+Lvysfiy8TgzMThy8TiztPlzNXmzev028Lgy8PgzMPhy9Dkx9XozeHuyP///xXNTX8AAB8KSURBVHja7N15dBzFnQdwzci6LGlkyYesy8i2jMeXfA2ObSQbX7KNhbGIA75vEDZYODbOcigOIAgk0bDWaR3WZUu+ZMknPoUJEASbACEh5NgcoJA4u5tdb9glm7DJTm33aGY0PdNdXVVd3TNd3fUHeS950eR98q3+/rq6R4oA5lK0IkwC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%3D%3D);");
                else
                    doc.Add("    background-image:url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAugAAAGyCAMAAACBecGuAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAxBpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6QkMwREUwMUZEQUQ1MTFFMzg0NzE4OTJDOTc2NzhENTQiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6QkMwREUwMUVEQUQ1MTFFMzg0NzE4OTJDOTc2NzhENTQiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiBXaW5kb3dzIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9IkM4MTkwMzM1QzJGQUNGREQyNjE1Nzc5QjVGQkY3QjUyIiBzdFJlZjpkb2N1bWVudElEPSJDODE5MDMzNUMyRkFDRkREMjYxNTc3OUI1RkJGN0I1MiIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PsYwArgAAAG5UExURf3j4/3k5fz8/OXl5e7u7v3l5vLy8v709P3m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);");
            }
            doc.Add("    /*padding-top: 23px;*/");
            doc.Add("}");
            doc.Add("</style>");
            doc.Add(Resources.style);

            doc.Add("</head>");
            doc.Add("<body>");
            doc.Add("<!-- header here -->");
        }

        protected void AddPage_End()
        {
            doc.Add(Resources.htmlEnd);
        }

        protected void AddWarning(string message)
        {
            doc.Add("<br><span class='Subtitle' style='color:red !important;'>" + message + "</span><br>");
            Log.n("Advarsel: " + message, null, true);
        }

        protected decimal Compute(DataTable table, string search, string filter = null)
        {
            object r = table.Compute(search, filter);
            if (!DBNull.Value.Equals(r))
                return Convert.ToDecimal(r);
            else
                return 0;
        }

        protected void OpenPage(string htmlFile)
        {
            if (!runningInBackground)
                browser.Url = new Uri(Path.Combine(FormMain.settingsPath, htmlFile));
        }
        protected void OpenPage_Loading()
        {
            if (!runningInBackground)
                browser.Url = new Uri(Path.Combine(FormMain.settingsPath, FormMain.htmlLoading));
        }

        protected void OpenPage_Stopped()
        {
            if (!runningInBackground)
                browser.Url = new Uri(Path.Combine(FormMain.settingsPath, FormMain.htmlStopped));
        }

        protected void OpenPage_Error()
        {
            if (!runningInBackground)
                browser.Url = new Uri(Path.Combine(FormMain.settingsPath, FormMain.htmlError));
        }

        protected void OpenPage_Import()
        {
            if (!runningInBackground)
                browser.Url = new Uri(Path.Combine(FormMain.settingsPath, FormMain.htmlImport));
        }
    }
}
