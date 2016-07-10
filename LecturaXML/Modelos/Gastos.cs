using System;

namespace LecturaXML.Modelos
{
    class Gastos : IEquatable<Gastos>, IComparable<Gastos>
    {        
        public short Id { get; set; }
        public string Fecha { get; set; }
        public string Concepto { get; set; }
        public string Folio { get; set; }
        public string Subtotal { get; set; }
        public string TotalImpuestosTrasladados { get; set; }
        public string IEPS { get; set; }        
        public string IVA { get; set; }
        public string Total { get; set; }
        public string Descripción { get; set; }
        public string Condiciones { get; set; }
        public string UUID { get; set; }

       public Gastos()
        {
            Concepto = "Sin Concepto";
            TotalImpuestosTrasladados = "Sin impuestos";
            IVA = "0.00";
            IEPS = "0.00";
            Condiciones = "Sin condiciones";
            Descripción = "";
        }

        public int CompareTo(Gastos other)
        {
            if (other == null)
            {
                return 1;
            }
            else
            {
                return this.Id.CompareTo(other.Id);
            }            
        }

        public bool Equals(Gastos other)
        {
            if (other == null)
            {
                return false;
            }
            return (this.Id.Equals(other.Id));
        }
    }
}
