namespace ControleFinanceiroAPI.Models;

public sealed class RegistrarProximosSalariosRequest
{
    public string Pessoa { get; set; } = string.Empty;
    public string UltimoMesAno { get; set; } = string.Empty;
    public decimal UltimoValorHora { get; set; }
    public decimal? NovoValorHora { get; set; }
    public List<MesSalarioRequest> Meses { get; set; } = new();
}

public sealed class MesSalarioRequest
{
    public string MesAno { get; set; } = string.Empty;
    public int HorasUteis { get; set; }
}
