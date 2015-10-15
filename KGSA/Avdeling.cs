using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class Avdeling
    {
        string[,] avd = new string[,]
        {
            {"1002", "Elkjøp Arendal"},
            {"1003", "Elkjøp Åsane"},
            {"1004", "Elkjøp Bodø"},
            {"1005", "Elkjøp fredrikstad"},
            {"1007", "Elkjøp Gjøvik"},
            {"1008", "Elkjøp Hamar"},
            {"1009", "Elkjøp Tønsberg"},
            {"1012", "Elkjøp Halden"},
            {"1013", "Elkjøp Sartor"},
            {"1014", "Elkjøp Kristiansand"},
            {"1016", "Elkjøp Lillehammer"},
            {"1017", "Elkjøp Lørenskog"},
            {"1018", "Elkjøp Sandvika"},
            {"1019", "Elkjøp Vinterbro"},
            {"1020", "Elkjøp Skien"},
            {"1023", "Elkjøp Tromsø"},
            {"1024", "Elkjøp Tiller"},
            {"1024", "Elkjøp Lade"},
            {"1025", "Elkjøp Ålesund"},
            {"1026", "Elkjøp Ullevål"},
            {"1028", "Elkjøp Fana"},
            {"1031", "Elkjøp Forus"},
            {"1032", "Elkjøp Kleppe"},
            {"1035", "Elkjøp Sarpsborg"},
            {"1036", "Elkjøp Sandefjord"},
            {"1037", "Elkjøp Larvik"},
            {"1038", "Elkjøp Molde"},
            {"1039", "Elkjøp Harstad"},
            {"1040", "Elkjøp Haugesund"},
            {"1042", "Elkjøp Jessheim"},
            {"1043", "Elkjøp Skøyen"},
            {"1044", "Elkjøp Lade"},
            {"1044", "Elkjøp Tiller"},
            {"1046", "Elkjøp Porsgrunn"},
            {"1047", "Elkjøp Kristiansund"},
            {"1049", "Elkjøp Kongsvinger"},
            {"1053", "Elkjøp Drammen"},
            {"1054", "Elkjøp Buskerud storsenter"},
            {"1055", "Elkjøp Notodden"},
            {"1091", "Elkjøp ECC Oslo"},
            {"1092", "Elkjøp Internett Oslo"},
            {"1201", "Elkjøp Åkra"},
            {"1207", "Elkjøp Dombås"},
            {"1215", "Elkjøp Florø"},
            {"1216", "Elkjøp Gloppen"},
            {"1217", "Elkjøp Hadeland"},
            {"1218", "Elkjøp Hønefoss"},
            {"1220", "Elkjøp Kongsberg"},
            {"1221", "Elkjøp Leknes"},
            {"1222", "Elkjøp Svolvær"},
            {"1224", "Elkjøp Lyngdal"},
            {"1226", "Elkjøp Mo i Rana"},
            {"1228", "Elkjøp Moss"},
            {"1229", "Elkjøp Mysen"},
            {"1235", "Elkjøp Namsos"},
            {"1236", "Elkjøp Rjukan/Tinn"},
            {"1238", "Elkjøp Levanger"},
            {"1240", "Elkjøp Steinkjer"},
            {"1241", "Elkjøp Stokmarknes"},
            {"1242", "Elkjøp Os"},
            {"1243", "Elkjøp Sortland"},
            {"1245", "Elkjøp Vinje"},
            {"1246", "Elkjøp Otta"},
            {"1247", "Elkjøp Voss"},
            {"1252", "Elkjøp Sogn"},
            {"1254", "Elkjøp Førde"},
            {"1258", "Elkjøp Askøy"},
            {"1261", "Elkjøp Bjergsted"},
            {"1262", "Elkjøp Stavanger Sentrum"},
            {"1263", "Elkjøp Horten"},
            {"1265", "Elkjøp Ski"},
            {"1267", "Elkjøp Husnes"},
            {"1271", "Elkjøp Ørsta"},
            {"1272", "Elkjøp Egersund"},
            {"1273", "Elkjøp Stord"},
            {"1275", "Elkjøp Fauske"},
            {"1276", "Elkjøp Nittedal"},
            {"1277", "Elkjøp Mosjøen"},
            {"1278", "Elkjøp Mo i Rana"},
            {"1300", "Lefdal Ulsteinvik"},
            {"1404", "Elkjøp Express Karl Johan"},
            {"1405", "Elkjøp Express Bogstad"},
            {"1409", "Elkjøp Express Storo"},
            {"1413", "Elkjøp Express Oasen"},
            {"1421", "Elkjøp Express Tveita"},
            {"1501", "Lefdal Sandvika"},
            {"1503", "Lefdal Strømmen"},
            {"1505", "Lefdal Lade"},
            {"1506", "Lefdal Alna"},
            {"1510", "Lefdal Tiller"},
            {"1512", "Lefdal Forus"},
            {"1513", "Lefdal Lagunen"},
            {"1514", "Lefdal Fredrikstad"},
            {"1515", "Lefdal Åsane"},
            {"1516", "Lefdal Drammen"},
            {"1517", "Lefdal Ålesund"},
            {"1518", "Lefdal Hamar"},
            {"1519", "Lefdal Kristiansand"},
            {"1520", "Lefdal Haugesund"},
            {"1521", "Lefdal Moss"},
            {"1522", "Lefdal Molde"},
            {"1524", "Lefdal Skien"},
            {"1525", "Lefdal Ski A.V"},
            {"1526", "Lefdal Storo"},
            {"1528", "Lefdal Tromsø"},
            {"1591", "Lefdal ECC Oslo"},
            {"1592", "Lefdal Internett Oslo"},
            {"1702", "Elkjøp Internett Kristiansand"},
            {"1703", "Elkjøp Internett Haugesund"},
            {"1704", "Elkjøp Internett Stavanger"},
            {"1705", "Elkjøp Internett Bergen"},
            {"1706", "Elkjøp Internett Trondheim"},
            {"1707", "Elkjøp Internett Larvik"},
            {"1708", "Elkjøp Internett Fauske"},
            {"1709", "Elkjøp Internett Tromsø"},
            {"1710", "Elkjøp Internett Molde"},
            {"1711", "Elkjøp Internett Ålesund"},
            {"1712", "Elkjøp Internett Førde"},
            {"1714", "Elkjøp Internett Hamar"},
            {"1727", "Lefdal Internett Kristiansand"},
            {"1729", "Lefdal Internett Stavanger"},
            {"1730", "Lefdal Internett Bergen"},
            {"1731", "Lefdal Internett Trondheim"},
            {"1732", "Lefdal Larvik Internett"},
            {"1733", "Lefdal Internett Fauske"},
            {"1734", "Lefdal Internett Tromsø"},
            {"1735", "Lefdal Internett Molde"},
            {"1737", "Lefdal Internett Førde"},
            {"1739", "Lefdal Internett Hamar"},
            {"1752", "Elkjøp ECC Kristiansand"},
            {"1753", "Elkjøp ECC Haugesund"},
            {"1754", "Elkjøp ECC Stavanger"},
            {"1755", "Elkjøp ECC Bergen"},
            {"1756", "Elkjøp ECC Trondheim"},
            {"1757", "Elkjøp Larvik ECC"},
            {"1758", "Elkjøp ECC Fauske"},
            {"1759", "Elkjøp ECC Tromsø"},
            {"1760", "Elkjøp ECC Molde"},
            {"1761", "Elkjøp ECC Ålesund"},
            {"1762", "Elkjøp ECC Førde"},
            {"1764", "Elkjøp ECC Hamar"},
            {"1777", "Lefdal ECC Kristiansand"},
            {"1778", "Lefdal ECC Haugesund"},
            {"1779", "Lefdal ECC Stavanger"},
            {"1780", "Lefdal ECC Bergen"},
            {"1781", "Lefdal ECC Trondheim"},
            {"1782", "Lefdal ECC Larvik"},
            {"1783", "Lefdal ECC Fauske"},
            {"1784", "Lefdal ECC Tromsø"},
            {"1785", "Lefdal ECC Molde"},
            {"1786", "Lefdal ECC Ålesund"},
            {"1787", "Lefdal ECC Førde"},
            {"1789", "Lefdal ECC Hamar"}
        };

        public string Get(int avdArg)
        {
            try
            {
                if (avdArg >= 1000 && avdArg <= 9999)
                {
                    for (int i = 0; i < avd.GetLength(0); i++)
                    {
                        if (avd[i, 0] == avdArg.ToString())
                            return avd[i, 1];
                    }
                    return avdArg.ToString();
                }
                return avdArg.ToString();
            }
            catch
            {
                return avdArg.ToString();
            }
        }

        public string Get(string avdArg)
        {
            try
            {
                int pass = Convert.ToInt32(avdArg);
                return Get(pass);
            }
            catch
            {
                return avdArg;
            }
        }
    }
}
