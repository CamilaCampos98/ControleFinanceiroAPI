namespace ControleFinanceiroAPI.Models;

public sealed class LoginResponse
{
    public bool Autenticado { get; set; }
    public string Usuario { get; set; } = string.Empty;
    public string Mensagem { get; set; } = string.Empty;
}
