namespace ControleFinanceiroAPI.Models
{
    public class EditarCompraRequest
    {
        public string IdLan { get; set; } = string.Empty;
        public string Compra { get; set; } = string.Empty;
        public decimal Valor { get; set; }
        public string FormaPgto { get; set; } = string.Empty;
        public string Cartao { get; set; } = string.Empty;
        public DateTime? Data { get; set; }

        public string Parcela { get; set; } = string.Empty;  // se quiser editar
        public string MesAno { get; set; } = string.Empty;   // se quiser editar
        public string Fonte { get; set; } = string.Empty;    // se quiser editar

        public string Pessoa { get; set; } = string.Empty;   // só se precisar mesmo
    }
}
