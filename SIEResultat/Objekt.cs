using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusResultat
{
    public class Objekt : IEquatable<Objekt>
    {
        private string typ;
        private string id;
        private string namn;

        public Objekt()
        {
            typ = "";
            id = "";
            namn = "";
        }

        public Objekt(string typ, string id, string namn)
        {
            this.typ = typ;
            this.id = id;
            this.namn = namn;
        }
        public static int Compare(Objekt o1, Objekt o2)
        {
            if (o1.Typ.Equals(o2.Typ))
                return o1.Id.CompareTo(o2.Id);
            else
                return o1.Typ.CompareTo(o2.Typ);
        }
        public int CompareTo(Objekt o)
        {
            if (o == null) return 1;
            if (this.Typ.Equals(o.Typ))
                return Id.CompareTo(o.Id);
            else
                return Typ.CompareTo(o.Typ);

        }

        public  bool Equals(Objekt o)
        {
            if (this.typ.Equals(o.Typ) & this.id.Equals(o.Id) & this.namn.Equals(o.Namn))
                return true;
            else
                return false;
        }

        public override bool Equals(Object o)
        {
            Objekt localO = (Objekt)o;
            if (this.typ.Equals(localO.Typ) & this.id.Equals(localO.Id) & this.namn.Equals(localO.Namn))
                return true;
            else
                return false;
        }



        public string Typ
        {
            get
            {
                return this.typ;
            }
            set
            {
                this.typ = value;
            }
        }
        public string Id
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }
        public string Namn
        {
            get
            {
                return this.namn;
            }
            set
            {
                this.namn = value;
            }
        }
    }
}
