namespace ControleFinanceiroAPI.Models
{
    public class CompraModel
    {
        public string FormaPgto { get; set; } = string.Empty; // "Crédito" ou "Débito"
        public int TotalParcelas { get; set; }  // 1 para débito
        public string Descricao { get; set; } = string.Empty; // Nome da compra
        public decimal ValorTotal { get; set; }
        public DateTime Data { get; set; }

    }
}
