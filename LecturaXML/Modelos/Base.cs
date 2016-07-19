using System;

namespace LecturaXML.Modelos
{
    class Base: IEquatable<Base>, IComparable<Base>
    {
        public short Id { get; set; }
        public string Fecha { get; set; }        
        public string Folio { get; set; }
        public string Serie { get; set; }
        public string Subtotal { get; set; }
        public string TotalImpuestosTrasladados { get; set; }
        public string IEPS { get; set; }
        public string IVA { get; set; }
        public string Total { get; set; }
        public string Descripción { get; set; }
        public string Condiciones { get; set; }
        public string UUID { get; set; }
        public string TipoComprobante { get; set; }

        public Base()
        {            
            TotalImpuestosTrasladados = "0.00";
            IVA = "0.00";
            IEPS = "0.00";            
            Descripción = "";
        }

        public int CompareTo(Base other)
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

        public bool Equals(Base other)
        {
            if (other == null)
            {
                return false;
            }
            return (this.Id.Equals(other.Id));
        }
    }
}
