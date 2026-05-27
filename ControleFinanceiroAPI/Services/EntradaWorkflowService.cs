using ControleFinanceiroAPI.Models;
using System.Globalization;

namespace ControleFinanceiroAPI.Services;

public sealed class EntradaWorkflowService
{
    private readonly GoogleSheetsService _googleSheetsService;
    private readonly ILogger<EntradaWorkflowService> _logger;

    public EntradaWorkflowService(GoogleSheetsService googleSheetsService, ILogger<EntradaWorkflowService> logger)
    {
        _googleSheetsService = googleSheetsService;
        _logger = logger;
    }

    public OperationResult RegistrarEntrada(EntradaModel? entrada)
    {
        if (entrada == null)
            return OperationResult.BadRequest("Dados inválidos.");

        try
        {
            if (!string.Equals(entrada.TipoEntrada, "Extra", StringComparison.OrdinalIgnoreCase))
            {
                return RegistrarSalario(entrada);
            }

            return RegistrarExtra(entrada);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao registrar entrada para {Pessoa}.", entrada.Pessoa);
            return OperationResult.Error($"Erro interno: {ex.Message}");
        }
    }

    private OperationResult RegistrarSalario(EntradaModel entrada)
    {
        var valorCalculado = entrada.ValorHora * entrada.HorasUteisMes;

        var linha = new List<object>
        {
            entrada.Pessoa,
            entrada.TipoEntrada,
            valorCalculado,
            entrada.MesAno,
            entrada.ValorHora,
            entrada.HorasExtras,
            string.Empty
        };

        _googleSheetsService.WriteEntrada(linha);

        return OperationResult.Ok(new
        {
            message = "Entrada registrada com sucesso!",
            valorCalculado,
            entrada
        });
    }

    private OperationResult RegistrarExtra(EntradaModel entrada)
    {
        var entradaBase = _googleSheetsService.GetEntradaPorPessoaEMes(entrada.Pessoa, entrada.MesAno);

        if (entradaBase == null)
            return OperationResult.BadRequest("Entrada base (salário) não encontrada para a pessoa e mês.");

        var valorHoraExtra = ParseValorPlanilha(entradaBase["ValorHora"]);
        if (valorHoraExtra == null)
            return OperationResult.BadRequest("Valor da hora inválido na base.");

        var extrasAtuais = ParseValorPlanilha(entradaBase["Extras"]);
        if (extrasAtuais == null)
            return OperationResult.BadRequest("Valor da extra inválido na base.");

        var valorExtraCalculado = valorHoraExtra.Value * entrada.HorasExtras;
        var novosExtras = valorExtraCalculado;

        _googleSheetsService.AtualizarExtrasEntrada(entrada.Pessoa, entrada.MesAno, novosExtras);

        return OperationResult.Ok(new
        {
            message = "Horas extras registradas com sucesso!",
            valorHoraExtra,
            horasExtras = entrada.HorasExtras,
            valorExtraCalculado,
            novosExtras
        });
    }

    private static decimal? ParseValorPlanilha(string valor)
    {
        var normalizado = valor
            .Replace(".", string.Empty)
            .Replace(",", ".");

        return decimal.TryParse(normalizado, NumberStyles.Any, CultureInfo.InvariantCulture, out var resultado)
            ? resultado
            : null;
    }
}
