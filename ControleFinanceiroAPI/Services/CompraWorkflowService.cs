using ControleFinanceiroAPI.Models;

namespace ControleFinanceiroAPI.Services;

public sealed class CompraWorkflowService
{
    private readonly GoogleSheetsService _googleSheetsService;
    private readonly ILogger<CompraWorkflowService> _logger;

    public CompraWorkflowService(GoogleSheetsService googleSheetsService, ILogger<CompraWorkflowService> logger)
    {
        _googleSheetsService = googleSheetsService;
        _logger = logger;
    }

    public OperationResult RegistrarCompra(CompraModel? compra)
    {
        if (compra == null)
            return OperationResult.BadRequest("Dados da compra não informados.");

        string mesFatura;
        try
        {
            mesFatura = _googleSheetsService.CalcularMesFatura(
                compra.Data,
                compra.Cartao,
                compra.Pessoa);
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogWarning(
                ex,
                "Não foi possível calcular a fatura da compra de {Pessoa} no cartão {Cartao}.",
                compra.Pessoa,
                compra.Cartao);
            return OperationResult.BadRequest($"Não foi possível calcular o período do cartão: {ex.Message}");
        }

        try
        {
            compra.MesAno = mesFatura;

            var temEntrada = _googleSheetsService.PessoaTemEntradaCadastrada(
                compra.Pessoa,
                mesFatura);

            if (!temEntrada)
            {
                return OperationResult.BadRequest(
                    $"A pessoa {compra.Pessoa} não possui entrada cadastrada no mês {mesFatura}.");
            }

            _googleSheetsService.WritePurchaseWithInstallments(compra);

            return OperationResult.Ok(new
            {
                id = compra.idLan,
                message = "Compra registrada com sucesso",
                mesAno = compra.MesAno
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao registrar compra para {Pessoa}.", compra.Pessoa);
            return OperationResult.Error($"Erro interno: {ex.Message}");
        }
    }
}
