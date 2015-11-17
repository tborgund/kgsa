using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public class KgsaTools
    {
        FormMain main;
        public KgsaTools(FormMain form) { main = form; }

        public string GetTempFilename(string shortName, string extension)
        {
            string filename = Path.ChangeExtension(shortName, extension);

            if (File.Exists(Path.Combine(FormMain.settingsTemp, filename)))
                filename = Path.ChangeExtension(shortName + " " + Path.GetRandomFileName(), extension);

            return Path.Combine(FormMain.settingsTemp, filename);
        }

        public string TextStyle_Shorten(string text, int maxChars)
        {
            try
            {
                if (text.Length > (maxChars + 2))
                    return text.Substring(0, maxChars) + "..";
            }
            catch (Exception) { };
            return text;
        }

        public string NumberStyle_Percent(object numeratorObj, object denominatorObj, bool colorNegative = false, bool colorPositive = false, decimal nullPoint = 1)
        {
            try
            {
                decimal numerator = 0;
                decimal denominator = 0;
                if (DBNull.Value == numeratorObj)
                    numerator = 0;
                else
                    numerator = Convert.ToDecimal(numeratorObj);
                if (DBNull.Value == denominatorObj)
                    denominator = 0;
                else
                    denominator = Convert.ToDecimal(denominatorObj);

                return NumberStyle_Percent(numerator, denominator, colorNegative, colorPositive, nullPoint);
            }
            catch (Exception) { }
            return "";
        }

        public string NumberStyle_Percent(decimal numerator, decimal denominator, bool colorNegative = false, bool colorPositive = false, decimal nullPoint = 1)
        {
            try
            {
                decimal value = 0;
                if (denominator != 0)
                    value = Math.Round(numerator / denominator, 10);

                string styleRed = "", styleGreen = "";
                if (colorNegative)
                    styleRed = " style='color:red'";
                if (colorPositive)
                    styleGreen = " style='color:green'";

                string format = "0.#";

                string output = main.appConfig.visningNull;
                if ((int)(value * 100) == 0)
                    return output;
                if (value < nullPoint)
                    output = "<span" + styleRed + ">" + (value * 100).ToString(format, FormMain.norway) + "%</span>";
                if (value >= nullPoint)
                    output = "<span" + styleGreen + ">" + (value * 100).ToString(format, FormMain.norway) + "%</span>";

                return output;
            }
            catch { }
            return main.appConfig.visningNull + "%";
        }

        public string NumberStyle_Normal(object obj, int decimals = 0, string afix = "", bool colorNegative = false, bool colorPositive = false)
        {
            try
            {
                if (DBNull.Value == obj)
                    return NumberStyle_Normal(0, decimals = 0, afix, colorNegative, colorPositive);
                else
                    return NumberStyle_Normal(Convert.ToDecimal(obj), decimals = 0, afix, colorNegative, colorPositive);
            }
            catch (Exception) { }
            return "";
        }

        public string NumberStyle_Normal(decimal value, int decimals = 0, string afix = "", bool colorNegative = false, bool colorPositive = false)
        {
            try
            {
                value = Math.Round(value, 10);

                string styleRed = "", styleGreen = "";
                if (colorNegative)
                    styleRed = " style='color:red'";
                if (colorPositive)
                    styleGreen = " style='color:green'";

                string format = "#,##0";
                if (decimals > 0)
                {
                    format += ".";
                    for (int i = 0; i < decimals; i++)
                        format += "#";
                }

                string output = main.appConfig.visningNull;
                if (value < 0)
                    output = "<span" + styleRed + ">" + value.ToString(format, FormMain.norway) + afix + "</span>";
                if (value > 0)
                    output = "<span" + styleGreen + ">" + value.ToString(format, FormMain.norway) + afix + "</span>";

                return output;
            }
            catch {}
            return main.appConfig.visningNull + afix;
        }

    }

    public static class UpdateServicePage
    {
        public static Form mainwin;
        public static event DelegateRunServiceList OnBrowseServicePage;
        public static event DelegateRunServiceSpecialList OnBrowseServiceList;

        public static void GetServicePage(string s, string t = "")
        {
            ThreadSafeBrowseServicePage(s, t);
        }

        public static void GetServiceList(string status, string filter)
        {
            ThreadSafeBrowseServiceList(status, filter);
        }

        private static void ThreadSafeBrowseServicePage(string s, string t = "")
        {
            if (mainwin != null && mainwin.InvokeRequired)
                mainwin.Invoke(new DelegateRunServiceList(ThreadSafeBrowseServicePage), new object[] { s, t });
            else
                OnBrowseServicePage(s, t);
        }

        private static void ThreadSafeBrowseServiceList(string status, string filter)
        {
            if (mainwin != null && mainwin.InvokeRequired)
                mainwin.Invoke(new DelegateRunServiceSpecialList(ThreadSafeBrowseServiceList), new object[] { status, filter });
            else
                OnBrowseServiceList(status, filter);
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;
            box.SelectionColor = color;
            box.AppendText(text);

            box.SelectionColor = box.ForeColor;

            box.Select(box.TextLength + 1 - text.Length, 22);
            box.SelectionColor = Color.Gray;
            box.SelectionLength = 0;
        }
    }

    public static class Test
    {
        public static IEnumerable<TSource> DistinctBy<TSource, TKey>
(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }
    }

    /// <summary>
    /// Moving average with 5 samples
    /// </summary>
    public class MovingAvg
    {
        private bool ready = false;
        private int cur = 0;
        private float[] a = new float[5] { 0, 0, 0, 0, 0 };

        public MovingAvg()
        {

        }

        public float Calc(float input)
        {
            if (input == 0)
                input = 0.00001F;

            if (a[0] == 0)
                a[0] = input;
            else if (a[1] == 0)
                a[1] = input;
            else if (a[2] == 0)
                a[2] = input;
            else if (a[3] == 0)
                a[3] = input;
            else if (a[4] == 0)
            {
                a[4] = input;
                ready = true;
            }
            else
            {
                a[cur] = input;

                if (cur == 0)
                    cur++;
                else if (cur == 1)
                    cur++;
                else if (cur == 2)
                    cur++;
                else if (cur == 3)
                    cur++;
                else if (cur == 4)
                    cur = 0;
            }

            float output = a.Average();
            if (ready)
                return output;
            else
                return input;
        }

        public void Clear()
        {
            a[0] = 0;
            a[1] = 0;
            a[2] = 0;
            a[3] = 0;
            a[4] = 0;
            a[5] = 0;
            cur = 0;
            ready = false;
        }
    }

    public static class StringExtensions
    {
        public static bool IsNullOrWhiteSpace(this string value)
        {
            if (value == null) return true;
            return string.IsNullOrEmpty(value.Trim());
        }
    }

    public enum Department { SDA, AudioVideo, MDA, Telecom, Computer, Kasse, Aftersales, Tekniker }
    public static class Store
    {
        public static string ToString(Department type)
        {
            if (type == Department.MDA)
                return "MDA";
            else if (type == Department.AudioVideo)
                return "AudioVideo";
            else if (type == Department.SDA)
                return "SDA";
            else if (type == Department.Telecom)
                return "Tele";
            else if (type == Department.Computer)
                return "Data";
            else if (type == Department.Kasse)
                return "Kasse";
            else if (type == Department.Aftersales)
                return "Aftersales";
            else if (type == Department.Tekniker)
                return "Tekniker";

            return "Invalid";
        }
    }
}
