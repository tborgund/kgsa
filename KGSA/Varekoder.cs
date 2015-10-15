using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Serialization;

namespace KGSA
{

    [XmlRoot("Varekode")]
    [XmlInclude(typeof(VarekodeList))] // include type class Person
    public class VarekodeList
    {
        [XmlElement("kode")]
        public string kode { get; set; }
        [XmlElement("provSelger")]
        public decimal provSelger { get; set; }
        [XmlElement("provTekniker")]
        public decimal provTekniker { get; set; }
        [XmlElement("salgspris")]
        public decimal salgspris { get; set; }
        [XmlElement("kategori")]
        public string kategori { get; set; }
        [XmlElement("synlig")]
        public bool synlig { get; set; }
        [XmlElement("inclhitrate")]
        public bool inclhitrate { get; set; }
        [XmlElement("alias")]
        public string alias { get; set; }
        public VarekodeList()
        {
            this.inclhitrate = true; // default value
        }

        public void Insert(string kodeArg, decimal provSelgerArg, decimal provTeknikerArg, decimal salgsprisArg, string kategoriArg, bool synligArg, string aliasArg, bool inchitrateArg)
        {
            this.kode = kodeArg;
            this.provSelger = provSelgerArg;
            this.provTekniker = provTeknikerArg;
            this.salgspris = salgsprisArg;
            this.kategori = kategoriArg;
            this.synlig = synligArg;
            this.alias = aliasArg;
            this.inclhitrate = inchitrateArg;
        }
    }
}
