namespace ControleFinanceiroAPI.Models
{
    public class EntradaModel
    {
        public string Pessoa { get; set; }
        public decimal ValorHora { get; set; }       // Valor da hora normal
        public decimal HorasUteisMes { get; set; }   // Horas normais trabalhadas no mês
        public string MesAno { get; set; }            // formato "MM/yyyy"
        public string TipoEntrada { get; set; }       // "Salario" ou "HoraExtra"
        public decimal HorasExtras { get; set; } = 0;     // Apenas para HoraExtra, quantidade de horas extras
    }

}
