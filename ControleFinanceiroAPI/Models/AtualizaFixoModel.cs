namespace ControleFinanceiroAPI.Models
{
    public class AtualizaFixoModel
    {
        public string Id { get; set; }
        public decimal Valor { get; set; }
        public bool Pago { get; set; }
        public bool Dividido { get; set; }
    }
}
