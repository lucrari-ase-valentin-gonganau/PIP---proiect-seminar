using System;

namespace ProiectIngineriaProgramarii.Models
{
    public class Produs
    {
        public int Id { get; set; }
        public string Nume { get; set; }
        public string Descriere { get; set; }
        public decimal Pret { get; set; }
        public int StocDisponibil { get; set; }
        public DateTime DataAdaugare { get; set; }
    }
}
