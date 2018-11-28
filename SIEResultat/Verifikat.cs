using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIEResultat
{
    
    class Verifikat
    {
        private string serie;
        private string verifikationsNummer;
        private string datum;
        private string verifikationsText;
        private string registreringsDatum;
        private string signatur;
        private List<Transaktion> transaktioner = new List<Transaktion>(); 

        public string Serie { get => serie; set => serie = value; }
        public string VerifikationsNummer { get => verifikationsNummer; set => verifikationsNummer = value; }
        public string Datum { get => datum; set => datum = value; }
        public string VerifikationsText { get => verifikationsText; set => verifikationsText = value; }
        public string RegistreringsDatum { get => registreringsDatum; set => registreringsDatum = value; }
        public string Signatur { get => signatur; set => signatur = value; }
        internal List<Transaktion> Transaktioner { get => transaktioner; set => transaktioner = value; }

        public Verifikat(string serie, string verifikationsNummer, string datum, string verifikationsText, string registreringsDatum="", string signatur="")
        {
            Serie = serie;
            VerifikationsNummer = verifikationsNummer;
            Datum = datum;
            VerifikationsText = verifikationsText;
            RegistreringsDatum = registreringsDatum;
            Signatur = signatur;
        }

        
    }
}
