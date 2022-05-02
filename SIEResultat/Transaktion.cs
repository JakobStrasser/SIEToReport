using System;
using System.Collections.Generic;

namespace EPPlusResultat
{
    class Transaktion
    {
        private Dictionary<string, string> objekt;
        private DateTime transaktionsdatum;
        private String kontonr;
        private double belopp;
        private string transtext;
        private double kvantitet;
        private string sign;


        public Dictionary<string, string> Objekt { get { if (objekt.Count == 0) objekt.Add("1", "Saknas"); return objekt; } set => objekt = value; }
        public DateTime Transaktionsdatum { get => transaktionsdatum; set => transaktionsdatum = value; }
        public string Kontonr { get => kontonr; set => kontonr = value; }
        public double Belopp { get => belopp; set => belopp = value; }
        public string Transtext { get => transtext; set => transtext = value; }
        public double Kvantitet { get => kvantitet; set => kvantitet = value; }
        public string Sign { get => sign; set => sign = value; }

        public Transaktion(Dictionary<string, string> objekt, DateTime transaktionsdatum, string kontonr, double belopp, string transtext, double kvantitet, string sign)
        {
            this.objekt = objekt;
            this.transaktionsdatum = transaktionsdatum;
            this.kontonr = kontonr;
            this.belopp = belopp;
            this.transtext = transtext;
            this.kvantitet = kvantitet;
            this.sign = sign;
        }


    }
}
