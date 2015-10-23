using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace KGSA
{
    //public class Pages
    //{
    //    FormMain main;
    //    public Page_Daily pageDaily;
    //    private static string PAGE_DAILY = "Daglig";
    //    public Pages(FormMain form)
    //    {
    //        main = form;
    //        pageDaily = new Page_Daily(main, PAGE_DAILY, main.appConfig.page_daily_time);
    //    }

    //}

    //public class Page
    //{
    //    protected FormMain main;
    //    public PageType type { get; }
    //    public string path { get; }
    //    public string name { get; }
    //    public DateTime updated { get; set; }
    //    public bool needRefresh
    //    {
    //        get
    //        {
    //            return File.Exists(path) && main.appConfig.page_refresh_time > updated;
    //        }
    //    }

    //    public Page(FormMain form, string pageName, DateTime updated)
    //    {
    //        main = form;
    //        name = pageName;
    //        path = Path.Combine(FormMain.settingsPath,"page_" +  name + ".html");
    //        this.updated = updated;
    //    }

    //    protected System.Windows.Forms.WebBrowser GetBrowser()
    //    {
    //        if (type == PageType.Ranking)
    //            return main.webRanking;
    //        if (type == PageType.Budget)
    //            return main.webBudget;
    //        if (type == PageType.Store)
    //            return main.webStore;
    //        if (type == PageType.Service)
    //            return main.webService;
    //        return null;
    //    }
    //}

    //public class Page_Daily : Page
    //{
    //    public Page_Daily(FormMain form, string pageName, DateTime updated) : base(form, pageName, updated) {}

    //    public void Create(bool runInBackground, BackgroundWorker worker, DateTime date, bool forceRefresh = false)
    //    {
    //        if (needRefresh || forceRefresh)
    //        {
    //            Logg.Debug("Siden " + name + " trengte å oppdateres");
    //            PageBudgetDaily page = new PageBudgetDaily(main, runInBackground, worker, GetBrowser());
    //            if (page.BuildPage(BudgetCategory.Daglig, main.appConfig.strBudgetDaily, FormMain.htmlBudgetDaily, date))
    //            {
    //                updated = DateTime.Now;
    //                main.appConfig.page_daily_time = updated;
    //            }
    //        }
    //        else
    //        {
    //            Logg.Debug("Siden " + name + " var allerede oppdatert - " + updated.ToShortDateString() + " : " + main.appConfig.page_refresh_time.ToShortDateString());
    //        }
    //    }

    //    public void Reset()
    //    {
    //        if (File.Exists(path))
    //            File.Delete(path);

    //        updated = FormMain.rangeMin;
    //    }
    //}

    //public enum PageType { Ranking, Budget, Store, Service }
}
