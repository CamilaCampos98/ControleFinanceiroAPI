namespace ControleFinanceiroAPI.Models
{
    public class CompraModel
    {
        public int idLan { get; set; }
        public string FormaPgto { get; set; } = string.Empty;
        public int TotalParcelas { get; set; }
        public string Descricao { get; set; } = string.Empty;
        public decimal ValorTotal { get; set; }
        public string? MesAno { get; set; }
        public DateTime Data { get; set; }
        public string Pessoa { get; set; } = string.Empty;  // ✅ Novo
        public string Fonte { get; set; } = string.Empty;   // ✅ Novo
    }
}
