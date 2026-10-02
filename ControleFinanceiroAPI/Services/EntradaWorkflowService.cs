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

    public OperationResult RegistrarProximosSalarios(RegistrarProximosSalariosRequest? request)
    {
        if (request == null || string.IsNullOrWhiteSpace(request.Pessoa) || request.Meses?.Count != 6)
            return OperationResult.BadRequest("Informe a pessoa e os seis períodos da prévia.");

        try
        {
            var linhas = _googleSheetsService.ReadData("Config!A1:L");
            var salarios = linhas.Skip(1)
                .Where(linha => string.Equals(linha.ElementAtOrDefault(0)?.ToString()?.Trim(),
                    request.Pessoa.Trim(), StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(linha.ElementAtOrDefault(1)?.ToString()?.Trim(), "Salario", StringComparison.OrdinalIgnoreCase))
                .Select(linha => new
                {
                    Linha = linha,
                    MesAno = linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? string.Empty,
                    PeriodoValido = DateTime.TryParseExact(linha.ElementAtOrDefault(3)?.ToString()?.Trim(),
                        "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var periodo),
                    Periodo = periodo
                })
                .Where(salario => salario.PeriodoValido)
                .OrderByDescending(salario => salario.Periodo)
                .FirstOrDefault();

            if (salarios == null)
                return OperationResult.BadRequest("Essa pessoa ainda não possui salário cadastrado.");

            if (!string.Equals(salarios.MesAno, request.UltimoMesAno, StringComparison.OrdinalIgnoreCase))
                return OperationResult.BadRequest("O último salário mudou desde a prévia. Atualize a página e confira novamente.");

            var valorHoraAtual = _googleSheetsService.ParseDecimal(salarios.Linha.ElementAtOrDefault(4)?.ToString());
            if (valorHoraAtual <= 0 || valorHoraAtual != request.UltimoValorHora)
                return OperationResult.BadRequest("O valor-hora do último salário mudou. Atualize a prévia antes de registrar.");

            if (request.NovoValorHora.HasValue && request.NovoValorHora.Value <= 0)
                return OperationResult.BadRequest("Informe um novo valor-hora maior que zero.");

            var valorHora = request.NovoValorHora ?? valorHoraAtual;
            var entradas = new List<IList<object>>();
            for (var indice = 0; indice < 6; indice++)
            {
                var mes = request.Meses[indice];
                var periodoEsperado = salarios.Periodo.AddMonths(indice + 1).ToString("MM/yyyy");
                if (!string.Equals(mes.MesAno, periodoEsperado, StringComparison.Ordinal) ||
                    mes.HorasUteis <= 0 || mes.HorasUteis % 8 != 0)
                    return OperationResult.BadRequest("A prévia contém períodos ou horas inválidos. Recalcule antes de registrar.");

                var linha = new List<object>
                {
                    salarios.Linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? request.Pessoa.Trim(),
                    "Salario",
                    valorHora * mes.HorasUteis,
                    periodoEsperado,
                    valorHora,
                    0
                };
                linha.AddRange(Enumerable.Range(6, 6)
                    .Select(coluna => salarios.Linha.ElementAtOrDefault(coluna) ?? string.Empty));
                entradas.Add(linha);
            }

            var primeiroPeriodo = salarios.Periodo.AddMonths(1);
            var linhasFixos = _googleSheetsService.ReadData("Fixos!A1:H");
            if (linhasFixos == null || linhasFixos.Count <= 1)
                return OperationResult.BadRequest("Essa pessoa não possui fixos anteriores para copiar.");
            var fixosDaPessoa = linhasFixos.Skip(1)
                .Where(linha => string.Equals(linha.ElementAtOrDefault(3)?.ToString()?.Trim(),
                    request.Pessoa.Trim(), StringComparison.OrdinalIgnoreCase))
                .Select(linha => new
                {
                    Linha = linha,
                    PeriodoValido = DateTime.TryParseExact(linha.ElementAtOrDefault(2)?.ToString()?.Trim(),
                        new[] { "MM/yyyy", "M/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None,
                        out var periodo),
                    Periodo = periodo
                })
                .Where(fixo => fixo.PeriodoValido)
                .ToList();

            var periodoOrigem = fixosDaPessoa
                .Where(fixo => fixo.Periodo < primeiroPeriodo)
                .Select(fixo => (DateTime?)fixo.Periodo)
                .Max();
            if (!periodoOrigem.HasValue)
                return OperationResult.BadRequest("Essa pessoa não possui fixos anteriores para copiar.");

            var fixosOrigem = fixosDaPessoa.Where(fixo => fixo.Periodo == periodoOrigem.Value).ToList();
            var fixoExistente = fixosDaPessoa.FirstOrDefault(fixo => fixo.Periodo >= primeiroPeriodo &&
                fixo.Periodo <= primeiroPeriodo.AddMonths(5));
            if (fixoExistente != null)
                return OperationResult.BadRequest($"Já existem fixos em {fixoExistente.Periodo:MM/yyyy}. " +
                    "Revise esse período antes de registrar o lote.");

            var quantidadeFixos = fixosOrigem.Count * 6;
            if (quantidadeFixos > 999)
                return OperationResult.BadRequest("Há fixos demais para copiar em uma única operação.");

            var novosFixos = new List<IList<object>>();
            var idBase = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() * 1000;
            for (var indice = 0; indice < 6; indice++)
            {
                var periodo = primeiroPeriodo.AddMonths(indice);
                var vencimento = new DateTime(periodo.AddMonths(1).Year, periodo.AddMonths(1).Month, 15);
                foreach (var fixo in fixosOrigem)
                {
                    novosFixos.Add(new List<object>
                    {
                        idBase + novosFixos.Count,
                        fixo.Linha.ElementAtOrDefault(1)?.ToString()?.Trim() ?? string.Empty,
                        periodo.ToString("MM/yyyy"),
                        salarios.Linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? request.Pessoa.Trim(),
                        vencimento,
                        _googleSheetsService.ParseDecimal(fixo.Linha.ElementAtOrDefault(5)?.ToString()),
                        "Não",
                        fixo.Linha.ElementAtOrDefault(7)?.ToString() ?? string.Empty
                    });
                }
            }

            _googleSheetsService.WriteEntradasEFixos(entradas, novosFixos);
            return OperationResult.Ok(new { quantidadeSalarios = entradas.Count, quantidadeFixos = novosFixos.Count,
                primeiroPeriodo = request.Meses[0].MesAno, ultimoPeriodo = request.Meses[5].MesAno,
                fixosOrigem = periodoOrigem.Value.ToString("MM/yyyy"), valorHora });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro ao registrar próximos salários de {Pessoa}.", request.Pessoa);
            return OperationResult.Error("Não foi possível registrar os próximos salários.");
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
