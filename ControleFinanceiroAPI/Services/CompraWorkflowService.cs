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

        try
        {
            var mesFatura = _googleSheetsService.CalcularMesFatura(
                compra.Data,
                compra.Cartao,
                compra.Pessoa);

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
                message = "Compra registrada com sucesso"
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao registrar compra para {Pessoa}.", compra.Pessoa);
            return OperationResult.Error($"Erro interno: {ex.Message}");
        }
    }
}
