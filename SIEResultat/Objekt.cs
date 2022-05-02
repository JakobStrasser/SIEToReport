using System;

namespace EPPlusResultat
{
    public class Objekt : IEquatable<Objekt> { 
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
            if (Typ.Equals(o.Typ))
                return Id.CompareTo(o.Id);
            else
                return Typ.CompareTo(o.Typ);

        }

        public bool Equals(Objekt o)
        {
            if (typ.Equals(o.Typ) & id.Equals(o.Id) & namn.Equals(o.Namn))
                return true;
            else
                return false;
        }

        public override bool Equals(Object o)
        {
            Objekt localO = (Objekt)o;
            if (typ.Equals(localO.Typ) & id.Equals(localO.Id) & namn.Equals(localO.Namn))
                return true;
            else
                return false;
        }



        public string Typ
        {
            get
            {
                return typ;
            }
            set
            {
                typ = value;
            }
        }
        public string Id
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }
        public string Namn
        {
            get
            {
                return namn;
            }
            set
            {
                namn = value;
            }
        }
    }
}
