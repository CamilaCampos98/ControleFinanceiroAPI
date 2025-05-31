namespace ControleFinanceiroAPI.Models
{
    public class DividirGastoModel
    {
        public string IdLinha { get; set; }
        public string NomeDestino { get; set; }
        public decimal ValorDividir { get; set; }
        public string Dividido { get; set; }
    }
}
