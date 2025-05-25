namespace ControleFinanceiroAPI.Models
{
    public class EntradaModel
    {
        public string Pessoa { get; set; } = string.Empty;
        public string Fonte { get; set; } = string.Empty;  // Ex.: Salário, Freelancer, etc.
        public decimal ValorHora { get; set; }
        public int HorasUteisMes { get; set; }
        public string MesAno { get; set; } = string.Empty; // Ex.: "05/2025"
    }

}
