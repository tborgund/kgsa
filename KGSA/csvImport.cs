using System;
using FileHelpers;
using System.Globalization;
using System.Threading;

namespace KGSA
{
    [IgnoreFirst(1)] 
    [DelimitedRecord(";")]
    public class csvImport
    {
        public string Avd;
        public string BilagsNr;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime Dato; 
        public string Sk;
        public string Varenummer;
        public string Beskrivelse;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal Antall;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal Salgspris;
        public string Kostpris;
        public string Btopercent;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal Btokr;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal Rab;
        public string Kgm;
        public string Merke;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal Mva;
        public string Kundenummer;
        public string Kundetype;
        public string Uke;

    }

    [IgnoreFirst(2)]
    [DelimitedRecord(";")]
    public class csvService
    {
        public int Avd;
        public int Ordrenr;
        public string Navn;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime DatoMottatt;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime? DatoIarbeid;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime? DatoFerdig;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime? DatoUtlevert;
        public int Dag;
        public string Status;
        public string Selgerkode;
        public string Verksted;
        [FieldConverter(ConverterKind.Date, "dd.MM.yy")]
        public DateTime LoggDato;
        [FieldConverter(ConverterKind.Date, "HH:mm")]
        public DateTime LoggTid;
        public string LoggKode;
        public string LoggTekst;
        public string Ekode;
    }

    [IgnoreFirst(1)]
    [DelimitedRecord(";")]
    public class csvObsolete
    {
        public static CultureInfo norway = new CultureInfo("nb-NO");

        public int Avd;
        public string AvdNavn;
        public int Kat;
        public string KatNavn;
        public int Grp;
        public string GrpNavn;
        public int Mod;
        public string ModNavn;
        public int Merke;
        public string MerkeNavn;
        public string Varekode;
        public string VareTekst;
        public int T;
        public string Referanse;
        public int AntallLager;
        [FieldConverter(ConverterKind.Decimal, ",")]
        public decimal KostVerdiLager;
        [FieldConverter(typeof(ConvertDate))]
        public DateTime DatoInnLager;
        public int MndUkurans;
        public string UkuransProsent;
        public string UkuransVerdi;
        public string Overflow;
    }

    internal class ConvertDate : ConverterBase
    {
        /// <summary>
        /// different forms for date separator : . or / or space
        /// </summary>
        /// <param name="from">the string format of date - first the day</param>
        /// <returns></returns>
        public override object StringToField(string s)
        {
            // Set current culture to a culture that uses "." as DateSeparator
            Thread.CurrentThread.CurrentCulture = new CultureInfo("nb-NO");

            // This is better - Culture.InvariantCulture uses / for the DateTimeFormatInfo.DateSeparator
            // and you clearly express the intent to use the invariant culture
            DateTime dt;
            if (DateTime.TryParseExact(s, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            else
                throw new ArgumentException("can not make a date from " + s, "from");
        }
    }
}
