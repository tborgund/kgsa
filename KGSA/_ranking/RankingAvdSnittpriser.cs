using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace KGSA
{
    class RankingAvdSnittpriser : Ranking
    {
        private DateTime dato;
        public RankingAvdSnittpriser() { }
        public RankingAvdSnittpriser(FormMain form, DateTime _dato)
        {
            this.main = form;
            this.dato = _dato;
            velgerPeriode = FormMain.datoPeriodeVelger;
        }

        private DataTable MakeTableDate(int rankMode, DateTime time)
        {
            try
            {
                string strMainVare = "";
                int[] intMainVare = main.appConfig.GetMainproductGroups(0);
                if (intMainVare.Length > 1)
                {
                    strMainVare = " AND (Varegruppe = " + intMainVare[0];
                    for (int i = 1; i < intMainVare.Length; i++)
                        strMainVare += " OR Varegruppe = " + intMainVare[i];
                    strMainVare += ")";
                }
                else if (intMainVare.Length == 1)
                    strMainVare = " AND Varegruppe = " + intMainVare[0];
                else
                    return ReadyTableSnittpris();

                DateTime from = time;
                DateTime to = time;
                DataTable sqlResult = new DataTable();

                if (rankMode == 0)
                {
                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }
                else if (rankMode == 1)
                {
                    from = GetFirstDayOfMonth(time);
                    to = GetLastDayOfMonth(time);

                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }
                else if (rankMode == 2) // bonus
                {
                    if (IsOdd(time.Month)) // Is at start of bonus periode
                        from = GetFirstDayOfMonth(time);
                    else // Is in second month of bonus periode, move one month
                        from = GetFirstDayOfMonth(time.AddMonths(-1));

                    if (IsOdd(time.Month)) // Is at start of bonus periode
                        to = GetLastDayOfMonth(time);
                    else // Is in second month of bonus periode, move one month
                        to = GetLastDayOfMonth(time.AddMonths(-1));

                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }
                else if (rankMode == 3)
                {
                    from = new DateTime(time.Year, 1, 1);
                    to = dtPick;

                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }
                else if (rankMode == 4)
                {
                    DateTime startOfLastWeek = main.database.GetStartOfLastWholeWeek(time); ;
                    from = startOfLastWeek;
                    to = startOfLastWeek.AddDays(6);

                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }
                else if (rankMode == 5)
                {
                    from = new DateTime(time.Year, 1, 1);
                    to = new DateTime(time.Year, 12, GetLastDayOfMonth(new DateTime(time.Year, 12, 1)).Day);

                    sqlResult = main.database.GetSqlDataTable("SELECT * FROM tblSalg WHERE (Dato >= '" + from.ToString("yyy-MM-dd") + "') AND (Dato <= '" + to.ToString("yyy-MM-dd") + "')" + strMainVare);
                }

                sqlResult.Columns.Add("SalgsprisExMva", typeof(Double), "Salgspris / Mva");

                DataTable dtWork = ReadyTableSnittpris();

                if (sqlResult.Rows.Count == 0)
                    return dtWork;

                if (main.arrayDbAvd == null || main.arrayDbAvd.ToList().Count == 0 && main.appConfig.rankingAvdelingShowAll)
                {
                    Logg.Log("Vent til avdelingslisten er oppdatert..", Color.Red);
                    return null;
                    // Starter bakgrunnsoppgave for oppdatering av avdelinger
                    //if (!main.bwHentAvdelinger.IsBusy)
                    //{
                    //    Logg.Log("Oppdaterer avdelings liste, vent litt..");
                    //    main.bwHentAvdelinger.RunWorkerAsync();
                    //}
                    //else
                    //{
                    //    int waitTime = 0;
                    //    while (main.bwHentAvdelinger.IsBusy && waitTime < 30)
                    //    {
                    //        waitTime++;
                    //        System.Threading.Thread.Sleep(50);
                    //    }
                    //}
                }

                if (StopRankingPending())
                    return dtWork;

                List<string> avdList = FormMain.Favoritter;

                if (main.appConfig.rankingAvdelingShowAll && main.arrayDbAvd != null)
                    if (main.arrayDbAvd.Length > 0)
                        avdList = main.arrayDbAvd.ToList();

                string[,] gruppe2d = new string[5, 3] { { "1", "6", "MDA" }, { "2", "11", "LoB" }, { "3", "16", "SDA" }, { "4", "21", "Tele" }, { "5", "26", "Data" } };

                // Butikker
                foreach (string avdStr in avdList)
                {
                    if (StopRankingPending())
                        return dtWork;

                    Logg.Status("Oppdaterer [Snittpriser].. " + avdeling.Get(avdStr));

                    var rowsGet = sqlResult.Select("(Avdeling = " + avdStr + ")");
                    using (DataTable sqlResultAvd = rowsGet.Any() ? rowsGet.CopyToDataTable() : sqlResult.Clone())
                    {
                        decimal totalt_antall = 0, totalt_snitt_inntjen = 0, totalt_snitt_omset = 0, totalt_snitt_omsetExMva = 0;

                        int[] mainproductsAllList = main.appConfig.GetMainproductGroups(0);
                        foreach (int grp in mainproductsAllList)
                        {
                            totalt_antall += Compute(sqlResultAvd, "Count(Antall)", "[Varegruppe] = " + grp);
                            totalt_snitt_inntjen += Compute(sqlResultAvd, "Sum(Btokr)", "[Varegruppe] = " + grp);
                            totalt_snitt_omset += Compute(sqlResultAvd, "Sum(Salgspris)", "[Varegruppe] = " + grp);
                            totalt_snitt_omsetExMva += Compute(sqlResultAvd, "Sum(SalgsprisExMva)", "[Varegruppe] = " + grp);
                        }

                        if (totalt_antall == 0)
                            totalt_antall = 1;

                        DataRow dtRow = dtWork.NewRow();
                        dtRow["Avdeling"] = avdStr;
                        dtRow["Totalt"] = totalt_antall;
                        dtRow["Totalt_snitt_inntjen"] = totalt_snitt_inntjen / totalt_antall;
                        dtRow["Totalt_snitt_omset"] = totalt_snitt_omset / totalt_antall;
                        dtRow["Totalt_snitt_omsetExMva"] = totalt_snitt_omsetExMva / totalt_antall;
                        if (totalt_snitt_omsetExMva != 0)
                            dtRow["Totalt_snitt_margin"] = Math.Round(totalt_snitt_inntjen / totalt_snitt_omsetExMva * 100, 2);
                        else
                            dtRow["Totalt_snitt_margin"] = 0;

                        // Grupper
                        for (int i = 0; i < gruppe2d.GetLength(0); i++)
                        {
                            // Snittpriser gruppe
                            decimal gruppe_antall = 0, gruppe_snitt_inntjen = 0, gruppe_snitt_omset = 0, gruppe_snitt_omsetExMva = 0;

                            var rows = sqlResultAvd.Select("(Varegruppe >= " + gruppe2d[i, 0] + "00 AND Varegruppe <= " + gruppe2d[i, 0] + "99) OR Varegruppe > 900");
                            using (DataTable sqlResultAvdGruppe = rows.Any() ? rows.CopyToDataTable() : sqlResult.Clone())
                            {
                                int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(Convert.ToInt32(gruppe2d[i, 0]));
                                foreach (int grp in mainproductsGrpList)
                                {
                                    gruppe_antall += Compute(sqlResultAvdGruppe, "Count(Antall)", "[Varegruppe] = " + grp);
                                    gruppe_snitt_inntjen += Compute(sqlResultAvdGruppe, "Sum(Btokr)", "[Varegruppe] = " + grp);
                                    gruppe_snitt_omset += Compute(sqlResultAvdGruppe, "Sum(Salgspris)", "[Varegruppe] = " + grp);
                                    gruppe_snitt_omsetExMva += Compute(sqlResultAvdGruppe, "Sum(SalgsprisExMva)", "[Varegruppe] = " + grp);
                                }
                            }
                            if (gruppe_antall == 0)
                                gruppe_antall = 1;

                            dtRow[gruppe2d[i, 2]] = gruppe_antall;
                            dtRow[gruppe2d[i, 2] + "_snitt_inntjen"] = gruppe_snitt_inntjen / gruppe_antall;
                            dtRow[gruppe2d[i, 2] + "_snitt_omset"] = gruppe_snitt_omset / gruppe_antall;
                            dtRow[gruppe2d[i, 2] + "_snitt_omsetExMva"] = gruppe_snitt_omsetExMva / gruppe_antall;
                            if (gruppe_snitt_omsetExMva != 0)
                                dtRow[gruppe2d[i, 2] + "_snitt_margin"] = Math.Round(gruppe_snitt_inntjen / gruppe_snitt_omsetExMva * 100, 2);
                            else
                                dtRow[gruppe2d[i, 2] + "_snitt_margin"] = 0;
                        }

                        dtWork.Rows.Add(dtRow);
                    }
                }

                if (dtWork.Rows.Count > 0)
                {
                    decimal antallRows = dtWork.Rows.Count;

                    DataRow row = dtWork.NewRow();
                    row["Avdeling"] = "Alle";
                    row["Totalt"] = dtWork.Compute("Sum(Totalt)", null);
                    row["Totalt_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(Totalt_snitt_inntjen)", null)) / antallRows;
                    row["Totalt_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(Totalt_snitt_omset)", null)) / antallRows;
                    row["Totalt_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(Totalt_snitt_omsetExMva)", null)) / antallRows;

                    decimal totalt_inntjen = Convert.ToDecimal(row["Totalt_snitt_inntjen"]);
                    decimal totalt_omsetExMva = Convert.ToDecimal(row["Totalt_snitt_omsetExMva"]);
                    if (totalt_omsetExMva != 0)
                        row["Totalt_snitt_margin"] = Math.Round(totalt_inntjen / totalt_omsetExMva * 100, 2);
                    else
                        row["Totalt_snitt_margin"] = 0;

                    // MDA
                    row["MDA"] = dtWork.Compute("Sum(MDA)", null);
                    row["MDA_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(MDA_snitt_inntjen)", null)) / antallRows;
                    row["MDA_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(MDA_snitt_omset)", null)) / antallRows;
                    row["MDA_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(MDA_snitt_omsetExMva)", null)) / antallRows;

                    decimal mda_snitt_inntjen = Convert.ToDecimal(row["MDA_snitt_inntjen"]);
                    decimal mda_snitt_omsetExMva = Convert.ToDecimal(row["MDA_snitt_omsetExMva"]);
                    if (mda_snitt_omsetExMva != 0)
                        row["MDA_snitt_margin"] = Math.Round(mda_snitt_inntjen / mda_snitt_omsetExMva * 100, 2);
                    else
                        row["MDA_snitt_margin"] = 0;

                    // LoB
                    row["LoB"] = dtWork.Compute("Sum(LoB)", null);
                    row["LoB_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(LoB_snitt_inntjen)", null)) / antallRows;
                    row["LoB_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(LoB_snitt_omset)", null)) / antallRows;
                    row["LoB_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(LoB_snitt_omsetExMva)", null)) / antallRows;

                    decimal lob_snitt_inntjen = Convert.ToDecimal(row["LoB_snitt_inntjen"]);
                    decimal lob_snitt_omsetExMva = Convert.ToDecimal(row["LoB_snitt_omsetExMva"]);
                    if (lob_snitt_omsetExMva != 0)
                        row["LoB_snitt_margin"] = Math.Round(lob_snitt_inntjen / lob_snitt_omsetExMva * 100, 2);
                    else
                        row["LoB_snitt_margin"] = 0;

                    // SDA
                    row["SDA"] = dtWork.Compute("Sum(SDA)", null);
                    row["SDA_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(SDA_snitt_inntjen)", null)) / antallRows;
                    row["SDA_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(SDA_snitt_omset)", null)) / antallRows;
                    row["SDA_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(SDA_snitt_omsetExMva)", null)) / antallRows;

                    decimal sda_snitt_inntjen = Convert.ToDecimal(row["SDA_snitt_inntjen"]);
                    decimal sda_snitt_omsetExMva = Convert.ToDecimal(row["SDA_snitt_omsetExMva"]);
                    if (sda_snitt_omsetExMva != 0)
                        row["SDA_snitt_margin"] = Math.Round(sda_snitt_inntjen / sda_snitt_omsetExMva * 100, 2);
                    else
                        row["SDA_snitt_margin"] = 0;

                    // Tele
                    row["Tele"] = dtWork.Compute("Sum(Tele)", null);
                    row["Tele_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(Tele_snitt_inntjen)", null)) / antallRows;
                    row["Tele_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(Tele_snitt_omset)", null)) / antallRows;
                    row["Tele_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(Tele_snitt_omsetExMva)", null)) / antallRows;

                    decimal tele_snitt_inntjen = Convert.ToDecimal(row["Tele_snitt_inntjen"]);
                    decimal tele_snitt_omsetExMva = Convert.ToDecimal(row["Tele_snitt_omsetExMva"]);
                    if (tele_snitt_omsetExMva != 0)
                        row["Tele_snitt_margin"] = Math.Round(tele_snitt_inntjen / tele_snitt_omsetExMva * 100, 2);
                    else
                        row["Tele_snitt_margin"] = 0;

                    // Tele
                    row["Data"] = dtWork.Compute("Sum(Data)", null);
                    row["Data_snitt_inntjen"] = Convert.ToDecimal(dtWork.Compute("Sum(Data_snitt_inntjen)", null)) / antallRows;
                    row["Data_snitt_omset"] = Convert.ToDecimal(dtWork.Compute("Sum(Data_snitt_omset)", null)) / antallRows;
                    row["Data_snitt_omsetExMva"] = Convert.ToDecimal(dtWork.Compute("Sum(Data_snitt_omsetExMva)", null)) / antallRows;

                    decimal data_snitt_inntjen = Convert.ToDecimal(row["Data_snitt_inntjen"]);
                    decimal data_snitt_omsetExMva = Convert.ToDecimal(row["Data_snitt_omsetExMva"]);
                    if (data_snitt_omsetExMva != 0)
                        row["Data_snitt_margin"] = Math.Round(data_snitt_inntjen / data_snitt_omsetExMva * 100, 2);
                    else
                        row["Data_snitt_margin"] = 0;

                    dtWork.Rows.Add(row);
                }

                return dtWork;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }

        public List<string> GetTableHtml(int rankMode)
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;
                var hashId = random.Next(999, 99999);
                string urlID = "linkx";

                DataTable dtNow = MakeTableDate(rankMode, dato);
                if (dtNow.Rows.Count == 0)
                {
                    doc.Add("<br><span class='Subtitle' style='color:red !important;'>Fant ingen transaksjoner for angitt periode.</span><br>");
                    return doc;
                }

                DataTable dtForrige = new DataTable();
                DataTable dtIfjor = new DataTable();
                string strForrige = "", strIfjor = "";

                if (rankMode == 0)
                {
                    strForrige = "Forrige dag";
                    strIfjor = "I fjor";
                    dtForrige = MakeTableDate(rankMode, dato.AddDays(-1));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-1));
                }
                else if (rankMode == 1)
                {
                    strForrige = "Sist MTD";
                    strIfjor = "I fjor MTD";
                    dtForrige = MakeTableDate(rankMode, dato.AddMonths(-1));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-1));
                }
                else if (rankMode == 2)
                {
                    strForrige = "For.bns.mnd.";
                    strIfjor = "I fjor";
                    dtForrige = MakeTableDate(rankMode, dato.AddMonths(-2));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-1));
                }
                else if (rankMode == 3)
                {
                    strForrige = "For.YTD";
                    strIfjor = "To år siden";
                    dtForrige = MakeTableDate(rankMode, dato.AddYears(-1));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-2));
                }
                else if (rankMode == 4)
                {
                    strForrige = "For.uke";
                    strIfjor = "I fjor uke";
                    dtForrige = MakeTableDate(rankMode, dato.AddDays(-7));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-1));
                }
                else if (rankMode == 5)
                {
                    strForrige = "Forrige år";
                    strIfjor = "To år siden";
                    dtForrige = MakeTableDate(rankMode, dato.AddYears(-1));
                    dtIfjor = MakeTableDate(rankMode, dato.AddYears(-2));
                }

                dt = ReadyTableVisning();
                

                for (int i = 0; i < dtNow.Rows.Count; i++ )
                {
                    DataRow row = dt.NewRow();
                    row["Butikk"] = dtNow.Rows[i]["Avdeling"];
                    row["Periode"] = "Nå";
                    row["Hovedprodukter solgt"] = dtNow.Rows[i]["Totalt"];
                    row["Snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Totalt_snitt_inntjen"]), 0);
                    row["Snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Totalt_snitt_omset"]), 0);
                    row["Snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Totalt_snitt_margin"]), 2);
                    row["MDA solgt"] = dtNow.Rows[i]["MDA"];
                    row["MDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["MDA_snitt_inntjen"]), 0);
                    row["MDA snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["MDA_snitt_omset"]), 0);
                    row["MDA snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["MDA_snitt_margin"]), 2);
                    row["LoB solgt"] = dtNow.Rows[i]["LoB"];
                    row["LoB snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["LoB_snitt_inntjen"]), 0);
                    row["LoB snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["LoB_snitt_omset"]), 0);
                    row["LoB snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["LoB_snitt_margin"]), 2);
                    row["SDA solgt"] = dtNow.Rows[i]["SDA"];
                    row["SDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["SDA_snitt_inntjen"]), 0);
                    row["SDA snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["SDA_snitt_omset"]), 0);
                    row["SDA snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["SDA_snitt_margin"]), 2);
                    row["Tele solgt"] = dtNow.Rows[i]["Tele"];
                    row["Tele snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Tele_snitt_inntjen"]), 0);
                    row["Tele snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Tele_snitt_omset"]), 0);
                    row["Tele snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Tele_snitt_margin"]), 2);
                    row["Data solgt"] = dtNow.Rows[i]["Data"];
                    row["Data snittpris inntjen"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Data_snitt_inntjen"]), 0);
                    row["Data snittpris omset"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Data_snitt_omset"]), 0);
                    row["Data snittpris margin"] = Math.Round(Convert.ToDecimal(dtNow.Rows[i]["Data_snitt_margin"]), 2);
                    dt.Rows.Add(row);

                    DataRow[] result = dtForrige.Select("Avdeling = '" + dtNow.Rows[i]["Avdeling"].ToString() + "'");
                    if (result.Count() == 0)
                        continue;

                    DataRow rowF = dt.NewRow();
                    rowF["Butikk"] = result[0]["Avdeling"];
                    rowF["Periode"] = strForrige;
                    rowF["Hovedprodukter solgt"] = result[0]["Totalt"];
                    rowF["Snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["Totalt_snitt_inntjen"]), 0);
                    rowF["Snittpris omset"] = Math.Round(Convert.ToDecimal(result[0]["Totalt_snitt_omset"]), 0);
                    rowF["Snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["Totalt_snitt_margin"]), 2);
                    rowF["MDA solgt"] = result[0]["MDA"];
                    rowF["MDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["MDA_snitt_inntjen"]), 0);
                    rowF["MDA snittpris omset"] = Math.Round(Convert.ToDecimal(result[0]["MDA_snitt_omset"]), 0);
                    rowF["MDA snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["MDA_snitt_margin"]), 2);
                    rowF["LoB solgt"] = result[0]["LoB"];
                    rowF["LoB snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["LoB_snitt_inntjen"]), 0);
                    rowF["LoB snittpris omset"] = Math.Round(Convert.ToDecimal(result[0]["LoB_snitt_omset"]), 0);
                    rowF["LoB snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["LoB_snitt_margin"]), 2);
                    rowF["SDA solgt"] = result[0]["SDA"];
                    rowF["SDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["SDA_snitt_inntjen"]), 0);
                    rowF["SDA snittpris omset"] = Math.Round(Convert.ToDecimal(result[0]["SDA_snitt_omset"]), 0);
                    rowF["SDA snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["SDA_snitt_margin"]), 2);
                    rowF["Tele solgt"] = result[0]["Tele"];
                    rowF["Tele snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["Tele_snitt_inntjen"]), 0);
                    rowF["Tele snittpris omset"] = Math.Round(Convert.ToDecimal(result[0]["Tele_snitt_omset"]), 0);
                    rowF["Tele snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["Tele_snitt_margin"]), 2);
                    rowF["Data solgt"] = result[0]["Data"];
                    rowF["Data snittpris inntjen"] = Math.Round(Convert.ToDecimal(result[0]["Data_snitt_inntjen"]), 0);
                    rowF["Data snittpris omset"] =Math.Round(Convert.ToDecimal( result[0]["Data_snitt_omset"]), 0);
                    rowF["Data snittpris margin"] = Math.Round(Convert.ToDecimal(result[0]["Data_snitt_margin"]), 2);
                    dt.Rows.Add(rowF);

                    DataRow[] resultI = dtIfjor.Select("Avdeling = '" + dtNow.Rows[i]["Avdeling"].ToString() + "'");
                    if (resultI.Count() == 0)
                        continue;

                    DataRow rowI = dt.NewRow();
                    rowI["Butikk"] = resultI[0]["Avdeling"];
                    rowI["Periode"] = strIfjor;
                    rowI["Hovedprodukter solgt"] = resultI[0]["Totalt"];
                    rowI["Snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["Totalt_snitt_inntjen"]), 0);
                    rowI["Snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["Totalt_snitt_omset"]), 0);
                    rowI["Snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["Totalt_snitt_margin"]), 2);
                    rowI["MDA solgt"] = resultI[0]["MDA"];
                    rowI["MDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["MDA_snitt_inntjen"]), 0);
                    rowI["MDA snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["MDA_snitt_omset"]), 0);
                    rowI["MDA snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["MDA_snitt_margin"]), 2);
                    rowI["LoB solgt"] = resultI[0]["LoB"];
                    rowI["LoB snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["LoB_snitt_inntjen"]), 0);
                    rowI["LoB snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["LoB_snitt_omset"]), 0);
                    rowI["LoB snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["LoB_snitt_margin"]), 2);
                    rowI["SDA solgt"] = resultI[0]["SDA"];
                    rowI["SDA snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["SDA_snitt_inntjen"]), 0);
                    rowI["SDA snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["SDA_snitt_omset"]), 0);
                    rowI["SDA snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["SDA_snitt_margin"]), 2);
                    rowI["Tele solgt"] = resultI[0]["Tele"];
                    rowI["Tele snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["Tele_snitt_inntjen"]), 0);
                    rowI["Tele snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["Tele_snitt_omset"]), 0);
                    rowI["Tele snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["Tele_snitt_margin"]), 2);
                    rowI["Data solgt"] = resultI[0]["Data"];
                    rowI["Data snittpris inntjen"] = Math.Round(Convert.ToDecimal(resultI[0]["Data_snitt_inntjen"]), 0);
                    rowI["Data snittpris omset"] = Math.Round(Convert.ToDecimal(resultI[0]["Data_snitt_omset"]), 0);
                    rowI["Data snittpris margin"] = Math.Round(Convert.ToDecimal(resultI[0]["Data_snitt_margin"]), 2);
                    dt.Rows.Add(rowI);
                }

                main.openXml.SaveDocument(dt, "Snittpriser", "Alle avdelinger", dato, "Snittpriser - Periode: " + main.GetRankModeText(main.appConfig.rankingAvdelingMode) + " " + dtPick.ToString("dddd d. MMMM yyyy", norway));

                doc.Add("<div class='no-break'>");
                doc.Add("<h3>Alle avdelinger</h3>");

                doc.Add("<div class='toolbox hidePdf'>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleTable(" + hashId + ");' href='#'>Vis / Skjul</a><br>");
                doc.Add("<a class='GuiButton hidePdf' onclick='toggleAll();' href='#'>Alle</a><br>");
                doc.Add("</div>");

                doc.Add("<table class='" + outerclass + " toggleAll' id='" + hashId + "'><tr><td>");
                doc.Add("<table class='tablesorter'>");
                doc.AddRange(MakeTableHeader());
                doc.Add("<tbody>");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string styleUnderline = "";
                    if (dt.Rows[i]["Periode"].ToString() == "Nå")
                        styleUnderline = "border-bottom: 1px solid #c78425;";
                    else if (dt.Rows[i]["Periode"].ToString() == strIfjor && dt.Rows[i]["Butikk"].ToString() != "Alle")
                        styleUnderline = "border-bottom: 1px solid black;";

                    if (dt.Rows[i]["Butikk"].ToString() != "Alle")
                    {
                        if (dt.Rows[i]["Periode"].ToString() == "Nå")
                            doc.Add("<tr><td class='text-cat' style='" + styleUnderline + "'><a href='#" + urlID + "b" + dt.Rows[i]["Butikk"].ToString() + "'>" + avdeling.Get(Convert.ToInt32(dt.Rows[i]["Butikk"])).Replace(" ", "&nbsp;") + "</a></td>");
                        else
                            doc.Add("<tr><td class='text-cat' style='" + styleUnderline + "'><span style='color:#bfbfbf'>" + avdeling.Get(Convert.ToInt32(dt.Rows[i]["Butikk"])).Replace(" ", "&nbsp;") + "</span></td>");
                    }
                    else
                    {
                        if (dt.Rows[i]["Periode"].ToString() == "Nå")
                            doc.Add("<tr><td class='text-cat' style='" + styleUnderline + "'>" + dt.Rows[i]["Butikk"] + "</a></td>");
                        else
                            doc.Add("<tr><td class='text-cat' style='" + styleUnderline + "'><span style='color:#bfbfbf'>Alle</span></td>");
                    }

                    if (dt.Rows[i]["Periode"].ToString() != "Nå")
                        doc.Add("<td class='text-cat' style='" + styleUnderline + "'>" + dt.Rows[i]["Periode"] + "</a></td>");
                    else
                        doc.Add("<td class='text-cat' style='" + styleUnderline + "'>&nbsp;</td>");

                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Hovedprodukter solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["Snittpris margin"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["MDA solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["MDA snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["MDA snittpris margin"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["LoB solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["LoB snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["LoB snittpris margin"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["SDA solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["SDA snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["SDA snittpris margin"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Tele solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Tele snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["Tele snittpris margin"].ToString()) + "</td>");

                    doc.Add("<td class='numbers-gen' style='border-left:2px solid #000;" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Data solgt"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-gen' style='" + styleUnderline + "'>" + PlusMinus(dt.Rows[i]["Data snittpris omset"].ToString()) + "</td>");
                    doc.Add("<td class='numbers-percent' style='" + styleUnderline + "'>" + PercentShare(dt.Rows[i]["Data snittpris margin"].ToString()) + "</td>");

                    doc.Add("</tr>");
                }
                doc.Add("</table></td></tr></table>");
                doc.Add("</div>");
                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return new List<string> { "Feil oppstod under ranking av snittpriser." };
            }
        }

        public List<string> GetVaregrupper()
        {
            try
            {
                var doc = new List<string>();
                if (StopRankingPending())
                    return doc;

                string strVaregrp = "";
                var dt = main.database.GetSqlDataTable("select distinct Varegruppe, VaregruppeNavn from tblVareinfo order by Varegruppe asc");
                if (dt == null)
                    return doc;

                if (dt.Rows.Count > 0)
                {
                    int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(0);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        foreach (int grp in mainproductsGrpList)
                        {
                            if (Convert.ToInt32(dt.Rows[i][0]) == grp)
                                strVaregrp += ", " + dt.Rows[i][0] + ": " + dt.Rows[i][1];
                        }
                    }
                    if (strVaregrp.Length > 3)
                        strVaregrp = strVaregrp.Substring(2, strVaregrp.Length - 2);

                    if (strVaregrp.Length > 3)
                    {
                        doc.Add("<h3>Hovedprodukt varegrupper</h3>");
                        doc.Add("<span class='Subtext'>" + strVaregrp + "</span>");
                    }
                }
                else
                {
                    int[] mainproductsGrpList = main.appConfig.GetMainproductGroups(0);
                    for (int i = 0; i < mainproductsGrpList.Length; i++)
                        strVaregrp += ", " + mainproductsGrpList[i];

                    if (strVaregrp.Length > 3)
                        strVaregrp = strVaregrp.Substring(2, strVaregrp.Length - 2);

                    if (strVaregrp.Length > 3)
                    {
                        doc.Add("<h3>Hovedprodukt varegrupper</h3>");
                        doc.Add("<span class='Subtext'>" + strVaregrp + "</span>");
                    }
                }

                return doc;
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
                return null;
            }
        }
        
        public List<string> MakeTableHeader()
        {
            List<string> doc = new List<string> { };

            doc.Add("<thead><tr>");

            doc.Add("<th class=\"{sorter: 'text'}\" width=95 >&nbsp;<br>Butikk</td>");
            doc.Add("<th class=\"{sorter: 'text'}\" width=35 >&nbsp;<br>&nbsp;</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >Snittpris<br>Solgt</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=30  style='border-left:2px solid #000;'>MDA<br>salg</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=30  style='border-left:2px solid #000;'>LoB<br>salg</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=30  style='border-left:2px solid #000;'>SDA<br>salg</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=30  style='border-left:2px solid #000;'>Tele<br>salg</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("<th class=\"{sorter: 'digit'}\" width=30  style='border-left:2px solid #000;'>Data<br>salg</td>");
            //doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Inntjen</td>");
            doc.Add("<th class=\"{sorter: 'digit'}\" width=45 >&nbsp;<br>Omset</td>");
            doc.Add("<th class=\"{sorter: 'procent'}\" width=35 >&nbsp;<br>Margin</td>");

            doc.Add("</tr></thead>");
            return doc;
        }


        public DataTable ReadyTableSnittpris()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Avdeling", typeof(string));
            dataTable.Columns.Add("Totalt", typeof(decimal));
            dataTable.Columns.Add("Totalt_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("Totalt_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("Totalt_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("Totalt_snitt_margin", typeof(decimal));
            dataTable.Columns.Add("MDA", typeof(decimal)); // 6
            dataTable.Columns.Add("MDA_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("MDA_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("MDA_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("MDA_snitt_margin", typeof(decimal));
            dataTable.Columns.Add("LoB", typeof(decimal)); // 11
            dataTable.Columns.Add("LoB_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("LoB_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("LoB_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("LoB_snitt_margin", typeof(decimal));
            dataTable.Columns.Add("SDA", typeof(decimal)); // 16
            dataTable.Columns.Add("SDA_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("SDA_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("SDA_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("SDA_snitt_margin", typeof(decimal));
            dataTable.Columns.Add("Tele", typeof(decimal));
            dataTable.Columns.Add("Tele_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("Tele_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("Tele_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("Tele_snitt_margin", typeof(decimal));
            dataTable.Columns.Add("Data", typeof(decimal));
            dataTable.Columns.Add("Data_snitt_inntjen", typeof(decimal));
            dataTable.Columns.Add("Data_snitt_omset", typeof(decimal));
            dataTable.Columns.Add("Data_snitt_omsetExMva", typeof(decimal));
            dataTable.Columns.Add("Data_snitt_margin", typeof(decimal));

            return dataTable;
        }

        public DataTable ReadyTableVisning()
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Butikk", typeof(string));
            dataTable.Columns.Add("Periode", typeof(string));
            dataTable.Columns.Add("Hovedprodukter solgt", typeof(decimal));
            dataTable.Columns.Add("Snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("Snittpris omset", typeof(decimal));
            dataTable.Columns.Add("Snittpris margin", typeof(decimal));
            dataTable.Columns.Add("MDA solgt", typeof(decimal)); // 6
            dataTable.Columns.Add("MDA snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("MDA snittpris omset", typeof(decimal));
            dataTable.Columns.Add("MDA snittpris margin", typeof(decimal));
            dataTable.Columns.Add("LoB solgt", typeof(decimal)); // 11
            dataTable.Columns.Add("LoB snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("LoB snittpris omset", typeof(decimal));
            dataTable.Columns.Add("LoB snittpris margin", typeof(decimal));
            dataTable.Columns.Add("SDA solgt", typeof(decimal)); // 16
            dataTable.Columns.Add("SDA snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("SDA snittpris omset", typeof(decimal));
            dataTable.Columns.Add("SDA snittpris margin", typeof(decimal));
            dataTable.Columns.Add("Tele solgt", typeof(decimal));
            dataTable.Columns.Add("Tele snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("Tele snittpris omset", typeof(decimal));
            dataTable.Columns.Add("Tele snittpris margin", typeof(decimal));
            dataTable.Columns.Add("Data solgt", typeof(decimal));
            dataTable.Columns.Add("Data snittpris inntjen", typeof(decimal));
            dataTable.Columns.Add("Data snittpris omset", typeof(decimal));
            dataTable.Columns.Add("Data snittpris margin", typeof(decimal));
            return dataTable;
        }
    }
}
