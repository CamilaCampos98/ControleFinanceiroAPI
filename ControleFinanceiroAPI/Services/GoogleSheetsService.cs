using ControleFinanceiroAPI.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;

public class GoogleSheetsService
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "ControleFinanceiro";

    private readonly string SpreadsheetId = "16c4P1KwZfuySZ36HSBKvzrl4ZagEXioD6yDhfQ9fhjM";
    private readonly string SheetName = "Controle";

    private readonly SheetsService _service;

    public GoogleSheetsService()
    {
        var jsonFilePath = GetJsonFilePath();
        if (!File.Exists(jsonFilePath))
        {
            throw new FileNotFoundException($"Arquivo de credenciais não encontrado em: {jsonFilePath}");
        }

        GoogleCredential credential;
        using (var stream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                                         .CreateScoped(Scopes);
        }

        _service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });
    }

    private static string GetJsonFilePath()
    {
        var pathFromEnv = Environment.GetEnvironmentVariable("GOOGLE_SHEETS_JSON_PATH");

        if (!string.IsNullOrEmpty(pathFromEnv) && File.Exists(pathFromEnv))
        {
            return pathFromEnv;
        }

        // Tenta na raiz
        var localPath = Path.Combine(Directory.GetCurrentDirectory(), "credentials.json");
        if (File.Exists(localPath))
        {
            return localPath;
        }

        // Tenta na pasta wwwroot
        var wwwrootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "credentials.json");
        if (File.Exists(wwwrootPath))
        {
            return wwwrootPath;
        }

        // Se não encontrar, retorna o primeiro caminho para lançar erro depois
        return pathFromEnv ?? localPath;
    }

    public IList<IList<object>> ReadData(string range)
    {
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        return response.Values;
    }

    public void WritePurchaseWithInstallments(CompraModel compra)
    {
        var idLan = ObterProximoIdLan();
        compra.idLan = idLan;

        // Tratamento da Fonte
        compra.Fonte = string.IsNullOrWhiteSpace(compra.Fonte) || compra.Fonte == "string"
            ? "Salário"
            : compra.Fonte.Trim();

        // Tratamento da Forma de Pagamento
        if (!string.IsNullOrWhiteSpace(compra.FormaPgto))
        {
            var forma = compra.FormaPgto.Trim().ToUpper();
            if (forma == "D")
                compra.FormaPgto = "Débito";
            else if (forma == "C")
                compra.FormaPgto = "Crédito";
            else
                compra.FormaPgto = forma; // Mantém o que veio, se não for D ou C
        }
        else
        {
            compra.FormaPgto = "Débito"; // Ou outro default se quiser
        }
        var linhas = new List<IList<object>>();
        var valorParcela = compra.ValorTotal / compra.TotalParcelas;

        for (int i = 1; i <= compra.TotalParcelas; i++)
        {
            var dataParcela = compra.Data.AddMonths(i - 1);
            var parcelaStr = compra.TotalParcelas > 1 ? $"{i}/{compra.TotalParcelas}" : "";

            linhas.Add(new List<object> {
            compra.idLan,
            compra.FormaPgto,
            parcelaStr,
            compra.Descricao,
            valorParcela,
            dataParcela.ToString("MM/yyyy"),
            dataParcela.ToString("yyyy-MM-dd"),
            compra.Pessoa,
            compra.Fonte
        });
        }

        var valueRange = new ValueRange { Values = linhas };

        var appendRequest = _service.Spreadsheets.Values.Append(
            valueRange,
            SpreadsheetId,
            $"{SheetName}"
        );
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        var response = appendRequest.Execute();

        // Opcional: log para confirmar inserção
        Console.WriteLine("Linhas inseridas: " + response.Updates.UpdatedRows);
    }


    public void TestConnection()
    {
        var request = _service.Spreadsheets.Get(SpreadsheetId);
        var response = request.Execute();
        Console.WriteLine($"Planilha encontrada: {response.Properties.Title}");
    }

    public int ObterProximoIdLan()
    {
        var range = $"{SheetName}!A:A"; // Supondo que idLan está na coluna A
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();

        var valores = response.Values;

        if (valores == null || valores.Count == 0)
            return 1; // Se não tem nenhum, começa com 1

        // Pular o cabeçalho se houver
        var idList = valores.Skip(1) // remove cabeçalho, se não tiver, tira esse skip
                             .Select(l => int.TryParse(l.FirstOrDefault()?.ToString(), out var id) ? id : 0)
                             .Where(id => id > 0)
                             .ToList();

        if (!idList.Any())
            return 1;

        return idList.Max() + 1;
    }

    public (bool Success, string Message, object? Data) GetResumoPorPessoaEPeriodo(string pessoa, DateTime inicio, DateTime fim)
    {
        try
        {
            var linhas = ReadData($"{SheetName}!A:I");

            if (linhas == null || linhas.Count <= 1)
                return (false, "Nenhum dado encontrado na planilha.", null);

            var header = linhas[0];

            // Leitura da aba Config para buscar o salário
            var config = ReadData("Config!A:D");

            decimal salario = 0;
            foreach (var linha in config.Skip(1)) // Ignora o header
            {
                var pessoaConfig = linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";
                var mesAno = linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";

                if (string.Equals(pessoaConfig, pessoa, StringComparison.OrdinalIgnoreCase) &&
                    (mesAno == $"{inicio:MM/yyyy}" || mesAno == $"{fim:MM/yyyy}"))
                {
                    var salarioStr = linha.ElementAtOrDefault(2)?.ToString();
                    salario = ParseDecimal(salarioStr);
                    break;
                }
            }

            if (salario == 0)
            {
                return (false, "Salário não encontrado na aba Config para essa pessoa e período.", null);
            }

            // Filtrar as compras
            var comprasFiltradas = new List<Dictionary<string, object>>();
            decimal totalGasto = 0;

            for (int i = 1; i < linhas.Count; i++)
            {
                var linha = linhas[i];

                var pessoaLinha = linha.ElementAtOrDefault(7)?.ToString()?.Trim() ?? "";
                var dataStr = linha.ElementAtOrDefault(6)?.ToString()?.Trim() ?? "";

                if (!string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!DateTime.TryParse(dataStr, out DateTime dataCompra))
                    continue;

                if (dataCompra >= inicio && dataCompra <= fim)
                {
                    var valorStr = linha.ElementAtOrDefault(4)?.ToString();
                    var valor = ParseDecimal(valorStr);
                    totalGasto += valor;

                    var compra = new Dictionary<string, object>();
                    for (int j = 0; j < header.Count; j++)
                    {
                        var chave = header[j]?.ToString() ?? $"Coluna{j}";
                        var valorCelula = linha.ElementAtOrDefault(j) ?? "";
                        compra[chave] = valorCelula;
                    }

                    comprasFiltradas.Add(compra);
                }
            }

            var saldoRestante = salario - totalGasto;

            var resultado = new
            {
                Pessoa = pessoa,
                Periodo = $"{inicio:dd/MM/yyyy} a {fim:dd/MM/yyyy}",
                Salario = salario,
                TotalGasto = totalGasto,
                SaldoRestante = saldoRestante,
                Compras = comprasFiltradas
            };

            return (true, "Sucesso", resultado);
        }
        catch (Exception ex)
        {
            return (false, $"Erro ao acessar a planilha: {ex.Message}", null);
        }
    }


    public decimal ParseDecimal(string? valorStr)
    {
        if (string.IsNullOrWhiteSpace(valorStr))
            return 0;

        valorStr = valorStr.Replace("R$", "")
                            .Replace(".", "")   // Remove separadores de milhar
                            .Replace(",", ".")  // Troca vírgula decimal por ponto
                            .Trim();

        if (decimal.TryParse(valorStr,
                              System.Globalization.NumberStyles.Any,
                              System.Globalization.CultureInfo.InvariantCulture,
                              out decimal valor))
        {
            return valor;
        }

        return 0;
    }


    public void WriteEntrada(List<object> entrada)
    {
        var valueRange = new ValueRange { Values = new List<IList<object>> { entrada } };

        var appendRequest = _service.Spreadsheets.Values.Append(
            valueRange, SpreadsheetId, "Config!A:D");

        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        appendRequest.Execute();
    }

    public bool PessoaTemEntradaCadastrada(string pessoa, string mesAno)
    {
        var range = "Config!A:D"; // A = Pessoa, D = MesAno
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        var valores = response.Values;

        if (valores == null || valores.Count <= 1)
            return false;

        foreach (var linha in valores.Skip(1)) // 🔥 Pula o header
        {
            var pessoaPlanilha = linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";
            var mesAnoPlanilha = linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";

            if (string.Equals(pessoaPlanilha, pessoa, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(mesAnoPlanilha, mesAno, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }


    public decimal ConsultarSaldo(string pessoa, string mesAno, string fonte)
    {
        // Ler a aba Config
        var entradasRange = "Config!A:D";
        var entradas = _service.Spreadsheets.Values.Get(SpreadsheetId, entradasRange).Execute().Values ?? new List<IList<object>>();

        decimal totalEntrada = entradas
            .Where(r => r.Count >= 4 &&
                        r[0].ToString() == mesAno &&
                        r[1].ToString() == pessoa &&
                        r[3].ToString() == fonte)
            .Sum(r => decimal.TryParse(r[2].ToString(), out var v) ? v : 0);

        // Ler a aba Controle
        var controleRange = "Controle!A:I";
        var controle = _service.Spreadsheets.Values.Get(SpreadsheetId, controleRange).Execute().Values ?? new List<IList<object>>();

        decimal totalSaida = controle
            .Where(r => r.Count >= 9 &&
                        r[5].ToString() == mesAno &&
                        r[7].ToString() == pessoa &&
                        r[8].ToString() == fonte)
            .Sum(r => decimal.TryParse(r[4].ToString(), out var v) ? v : 0);

        return totalEntrada - totalSaida;
    }

}
