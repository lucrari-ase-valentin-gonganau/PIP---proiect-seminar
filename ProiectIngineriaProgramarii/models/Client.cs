using System;

namespace ProiectIngineriaProgramarii.Models
{
    public class Client
    {
        public int Id { get; set; }
        public string Nume { get; set; }
        public string Prenume { get; set; }
        public string Email { get; set; }
        public string Telefon { get; set; }
        public string Adresa { get; set; }
        public DateTime DataInregistrare { get; set; }
    }
}
