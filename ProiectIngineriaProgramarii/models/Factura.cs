using System;
using System.Collections.Generic;

namespace ProiectIngineriaProgramarii.Models
{
    public class Factura
    {
        public int Id { get; set; }
        public string NumarFactura { get; set; }
        public DateTime DataEmitere { get; set; }
        public int ClientId { get; set; }
        public Client Client { get; set; }
        public List<ItemFactura> Itemi { get; set; }
        public decimal Subtotal { get; set; }
        public decimal TVA { get; set; }
        public decimal Total { get; set; }
        public string Observatii { get; set; }
        public string Status { get; set; }

        public Factura()
        {
            Itemi = new List<ItemFactura>();
            DataEmitere = DateTime.Now;
            Status = "Emisa";
        }

        public void CalculeazaTotaluri()
        {
            Subtotal = 0;
            foreach (var item in Itemi)
            {
                Subtotal += item.Subtotal;
            }
            TVA = Subtotal * 0.21m;
            Total = Subtotal + TVA;
        }
    }
}
