namespace ControleFinanceiroAPI.Models;

public sealed class SalarioCadastradoDTO
{
    public string Pessoa { get; init; } = string.Empty;
    public string MesAno { get; init; } = string.Empty;
    public decimal Valor { get; init; }
    public decimal ValorHora { get; init; }
    public decimal Extras { get; init; }
}
