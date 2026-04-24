namespace ProiectIngineriaProgramarii.Models
{
    public class ItemFactura
    {
        public int Id { get; set; }
        public int FacturaId { get; set; }
        public int ProdusId { get; set; }
        public string NumeProdus { get; set; }
        public int Cantitate { get; set; }
        public decimal PretUnitar { get; set; }
        public decimal Subtotal { get; set; }
        public string UnitateMasura { get; set; } = "buc";

        public void CalculeazaSubtotal()
        {
            Subtotal = Cantitate * PretUnitar;
        }
    }
}
