namespace ControleFinanceiroAPI.Models
{
    public class LinhaGastoModel
    {
        public long Id { get; set; }
        public string Pessoa { get; set; }
        public string Tipo { get; set; }
        public string MesAno { get; set; }
        public string Vencimento { get; set; }
        public decimal Valor { get; set; }
        public bool Pago { get; set; }
        public string Dividido { get; set; }
    }
}
