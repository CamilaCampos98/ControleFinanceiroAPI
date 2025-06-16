using ControleFinanceiroAPI.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

public class GoogleSheetsService
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "ControleFinanceiro";

    private readonly string SpreadsheetId = "16c4P1KwZfuySZ36HSBKvzrl4ZagEXioD6yDhfQ9fhjM";
    private readonly string SheetName = "Controle";
    private readonly string rangeFixo = "Fixos!A:H";
    private readonly string CartoesSheet = "Cartoes";
    private readonly string FixosTipoSheet = "TiposFixos";

    private readonly SheetsService _service;

    public GoogleSheetsService()
    {
        var credential = GetGoogleCredential().CreateScoped(Scopes);

        _service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });
    }

    private static GoogleCredential GetGoogleCredential()
    {
        // Primeiro tenta pegar das variáveis de ambiente (Render ou outro servidor)
        var credentialsJson = Environment.GetEnvironmentVariable("GOOGLE_CREDENTIALS");

        if (!string.IsNullOrEmpty(credentialsJson))
        {
            return GoogleCredential.FromJson(credentialsJson);
        }

        // Caso não encontre na variável, tenta os arquivos locais
        var possiblePaths = new[]
        {
        Environment.GetEnvironmentVariable("GOOGLE_SHEETS_JSON_PATH"), // caminho manual via env
        Path.Combine(Directory.GetCurrentDirectory(), "credentials.json"),
        Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "credentials.json")
        };

        foreach (var path in possiblePaths)
        {
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                return GoogleCredential.FromFile(path);
            }
        }

        throw new FileNotFoundException("Credenciais do Google não encontradas. Configure a variável de ambiente 'GOOGLE_SHEETS_CREDENTIALS_JSON' ou coloque o arquivo 'credentials.json' na raiz ou em wwwroot.");
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
                        compra.Fonte,
                        compra.Cartao
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
            var linhas = ReadData($"{SheetName}!A:J");

            if (linhas == null || linhas.Count <= 1)
                return (false, "Nenhum dado encontrado na planilha.", null);

            var header = linhas[0];

            // Leitura da aba Config para buscar o salário e extras
            var config = ReadData("Config!A:F");

            decimal salario = 0m;
            decimal extras = 0m;

            var competencia = ObterMesAnoCompetencia(inicio);

            foreach (var linha in config.Skip(1))
            {
                var pessoaConfig = linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";
                var mesAno = linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";

                if (string.Equals(pessoaConfig, pessoa, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(mesAno, competencia, StringComparison.OrdinalIgnoreCase))
                {
                    salario = ParseDecimal(linha.ElementAtOrDefault(2)?.ToString());
                    extras = ParseDecimal(linha.ElementAtOrDefault(5)?.ToString());
                    break;
                }
            }

            if (salario == 0)
            {
                return (false, "Salário não encontrado na aba Config para essa pessoa e período.", null);
            }

            decimal totalRecebido = salario + extras;

            var comprasFiltradas = new List<Dictionary<string, object>>();
            var resumoPorCartao = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var resumoPorCartaoTipo = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            decimal totalGasto = 0;

            for (int i = 1; i < linhas.Count; i++)
            {
                var linha = linhas[i];

                var pessoaLinha = linha.ElementAtOrDefault(7)?.ToString()?.Trim() ?? "";
                var dataStr = linha.ElementAtOrDefault(6)?.ToString()?.Trim() ?? "";
                var cartao = linha.ElementAtOrDefault(9)?.ToString()?.Trim() ?? "";

                if (!DateTime.TryParse(dataStr, out DateTime dataCompra))
                    continue;

                if (string.IsNullOrEmpty(cartao))
                    continue;

                string cartaoBase = RemoverPalavras(cartao);

                // Verifica se a compra está dentro do filtro da tela
                if (dataCompra >= inicio && dataCompra <= fim)
                {
                    var valor = ParseDecimal(linha.ElementAtOrDefault(4)?.ToString());

                    // Verifica se a compra é da pessoa
                    if (string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                    {
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

                    // ----------- Agrupamento por cartão considerando o ciclo específico -----------
                    var (inicioCorte, fimCorte) = ObterPeriodoCortePorCartao(dataCompra, cartaoBase);

                    if (dataCompra >= inicioCorte && dataCompra <= fimCorte)
                    {
                        // Se for cartão Bradesco, soma apenas se for da pessoa
                        if (cartaoBase.Equals("Bradesco", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }
                        }

                        // Soma no total do cartão
                        if (resumoPorCartao.ContainsKey(cartaoBase))
                            resumoPorCartao[cartaoBase] += valor;
                        else
                            resumoPorCartao[cartaoBase] = valor;

                        // Agrupa por tipo (Titular ou Adicional)
                        string tipo = cartao.Contains("Adicional", StringComparison.OrdinalIgnoreCase) ? "Adicional" :
                                      cartao.Contains("Titular", StringComparison.OrdinalIgnoreCase) ? "Titular" : "Outro";

                        string chave = $"{cartaoBase}|{tipo}";

                        if (resumoPorCartaoTipo.ContainsKey(chave))
                            resumoPorCartaoTipo[chave] += valor;
                        else
                            resumoPorCartaoTipo[chave] = valor;
                    }
                }
            }

            decimal totalGastoFixos = 0;

            // Ler dados da aba Fixos
            var linhasFixos = ReadData(rangeFixo);

            if (linhasFixos != null && linhasFixos.Count > 1)
            {
                for (int i = 1; i < linhasFixos.Count; i++)
                {
                    var linhaFixo = linhasFixos[i];

                    var pessoaFixo = linhaFixo.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";
                    var mesAnoFixo = linhaFixo.ElementAtOrDefault(2)?.ToString()?.Trim() ?? "";

                    if (!string.Equals(pessoaFixo, pessoa, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (mesAnoFixo == competencia)
                    {
                        var valorFixo = ParseDecimal(linhaFixo.ElementAtOrDefault(5)?.ToString());
                        totalGasto += valorFixo;
                        totalGastoFixos += valorFixo;
                    }
                }
            }

            var listaCartaoTipo = resumoPorCartaoTipo.Select(item =>
            {
                var partes = item.Key.Split('|');
                return new CartaoTipoResumo
                {
                    Cartao = partes.ElementAtOrDefault(0) ?? "Desconhecido",
                    Tipo = partes.ElementAtOrDefault(1) ?? "Desconhecido",
                    Valor = item.Value
                };
            }).OrderBy(x => x.Cartao).ToList();

            var saldoRestante = totalRecebido - totalGasto;

            var resultado = new
            {
                Pessoa = pessoa,
                Periodo = $"{inicio:dd/MM/yyyy} a {fim:dd/MM/yyyy}",
                Salario = salario,
                Extras = extras,
                TotalRecebido = totalRecebido,
                TotalGasto = totalGasto,
                SaldoRestante = saldoRestante,
                Compras = comprasFiltradas,
                GastosFixos = totalGastoFixos,
                resumoPorCartao = resumoPorCartao,
                resumoPorCartaoTipo = listaCartaoTipo
            };

            return (true, "Sucesso", resultado);
        }
        catch (Exception ex)
        {
            return (false, $"Erro ao acessar a planilha: {ex.Message}", null);
        }
    }


    public (bool Success, string Message, List<ResumoPessoaDTO>? Data) ResumoGeral()
    {
        try
        {
            var linhas = ReadData($"{SheetName}!A:J");
            if (linhas == null || linhas.Count <= 1)
                return (false, "Nenhum dado encontrado na planilha.", null);

            var config = ReadData("Config!A:F");
            if (config == null || config.Count <= 1)
                return (false, "Nenhum dado encontrado na aba Config.", null);

            var linhasFixos = ReadData(rangeFixo);

            var pessoas = config.Skip(1)
                .Select(l => l.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "")
                .Where(p => !string.IsNullOrEmpty(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // 📅 Definir mês vigente baseado na data atual e regra do dia 8
            DateTime hoje = DateTime.Today;
            DateTime mesVigente;

            if (hoje.Day <= 8)
                mesVigente = hoje.AddMonths(-1);
            else
                mesVigente = hoje;

            int mes = mesVigente.Month;
            int ano = mesVigente.Year;

            var resultado = new List<ResumoPessoaDTO>();

            foreach (var pessoa in pessoas)
            {
                // 🔍 Buscar configs da pessoa no mês vigente
                var linhaSalario = config.Skip(1)
                    .FirstOrDefault(l =>
                    {
                        var pessoaLinha = l.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";
                        var mesAnoStr = l.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";

                        if (!string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                            return false;

                        if (!DateTime.TryParseExact(mesAnoStr, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoData))
                            return false;

                        return mesAnoData.Month == mes && mesAnoData.Year == ano;
                    });

                if (linhaSalario == null)
                    continue; // Se não achou, ignora a pessoa

                decimal salario = ParseDecimal(linhaSalario?.ElementAtOrDefault(2)?.ToString());
                decimal extras = ParseDecimal(linhaSalario?.ElementAtOrDefault(5)?.ToString());

                decimal totalRecebido = salario + extras;

                // 📆 Intervalo do mês vigente
                var inicioMes = new DateTime(ano, mes, 1);
                var fimMes = inicioMes.AddMonths(1).AddDays(-1);

                // 🛒 Compras
                var comprasPessoa = linhas.Skip(1)
                    .Where(l =>
                    {
                        var dataStr = l.ElementAtOrDefault(6)?.ToString()?.Trim() ?? "";
                        var pessoaLinha = l.ElementAtOrDefault(7)?.ToString()?.Trim() ?? "";

                        if (!DateTime.TryParse(dataStr, out DateTime dataCompra))
                            return false;

                        return string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase) &&
                               dataCompra >= inicioMes && dataCompra <= fimMes;
                    })
                    .Select(l => new
                    {
                        Data = DateTime.TryParse(l.ElementAtOrDefault(6)?.ToString()?.Trim(), out DateTime d) ? d : (DateTime?)null,
                        Descricao = l.ElementAtOrDefault(3)?.ToString() ?? "",
                        Valor = ParseDecimal(l.ElementAtOrDefault(4)?.ToString())
                    })
                    .ToList();

                decimal totalGastosCompras = comprasPessoa.Sum(c => c.Valor);

                // 🏠 Fixos
                decimal totalGastosFixos = 0;
                if (linhasFixos != null && linhasFixos.Count > 1)
                {
                    totalGastosFixos = linhasFixos.Skip(1)
                        .Where(l =>
                        {
                            var pessoaLinha = l.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";
                            var mesAnoFixo = l.ElementAtOrDefault(2)?.ToString()?.Trim() ?? "";

                            if (!string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                                return false;

                            if (!DateTime.TryParseExact(mesAnoFixo, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoData))
                                return false;

                            return mesAnoData.Year == ano && mesAnoData.Month == mes;
                        })
                        .Sum(l => ParseDecimal(l.ElementAtOrDefault(5)?.ToString()));
                }

                // 🧮 Calcular saldo
                decimal totalGastos = totalGastosCompras + totalGastosFixos;
                decimal saldoRestante = totalRecebido - totalGastos;

                // 📅 Última compra
                var ultimaCompra = comprasPessoa.OrderByDescending(c => c.Data)
                    .FirstOrDefault();

                resultado.Add(new ResumoPessoaDTO
                {
                    Pessoa = pessoa,
                    SaldoRestante = saldoRestante,
                    UltimaCompra = ultimaCompra?.Data?.ToString("dd/MM/yyyy") ?? "-",
                    DescricaoUltimaCompra = ultimaCompra?.Descricao ?? "-"
                });
            }

            return (true, "Sucesso", resultado.OrderBy(x => x.Pessoa).ToList());
        }
        catch (Exception ex)
        {
            return (false, $"Erro: {ex.Message}", null);
        }
    }



    public (bool success, string message, object data) GetResumoPorPessoaEPeriodo(string pessoa, string mesAno)
    {
        try
        {

            // Ler dados das planilhas
            var configData = ReadData("Config!A1:F");
            var fixosData = ReadData("Fixos!A1:H");
            var controleData = ReadData("Controle!A1:J");

            // Mapear Config
            var configs = configData.Skip(1).Select(row => new
            {
                Pessoa = row[0].ToString(),
                Fonte = row[1].ToString(),
                Valor = ParseDecimal((row[2].ToString())),
                mesAno = row[3].ToString(),
                ValorHora = decimal.Parse(row[4].ToString()),
                Extras = decimal.Parse(row[5].ToString())
            }).ToList();

            // Mapear Fixos
            var fixos = fixosData.Skip(1).Select(row => new
            {
                Id = row[0].ToString(),
                Tipo = row[1].ToString(),
                mesAno = row[2].ToString(),
                Pessoa = row[3].ToString(),
                Vencimento = row[4].ToString(),
                Valor = ParseDecimal(row[5].ToString()),
                Pago = row[6].ToString(),
                Dividido = row[7].ToString()
            }).ToList();

            // Mapear Controle
            var controle = controleData.Skip(1).Select(row => new
            {
                IdLan = row[0].ToString(),
                FormaPgto = row[1].ToString(),
                Parcela = row[2].ToString(),
                Compra = row[3].ToString(),
                Valor = ParseDecimal(row[4].ToString()),
                mesAno = row[5].ToString(),
                Data = DateTime.Parse(row[6].ToString()),
                Pessoa = row[7].ToString(),
                Fonte = row[8].ToString(),
                Cartao = row[9].ToString()
            }).ToList();

            var dadosConfig = configs
                             .Where(c => c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && c.mesAno == mesAno)
                             .ToList();

            // Mesmo que não haja dados, consideramos salário e extras como 0
            var salario = dadosConfig.Any()
                ? dadosConfig.Where(c => c.Fonte.Equals("Salario", StringComparison.OrdinalIgnoreCase)).Sum(c => c.Valor)
                : 0;

            var extra = dadosConfig.Any()
                ? dadosConfig.Sum(c => c.Extras)
                : 0;

            // Filtrar fixos para pessoa e mesAno (mesAno convertido para string)
            var fixosPessoa = fixos
                .Where(f => f.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && f.mesAno == mesAno.ToString())
                .Sum(f => f.Valor);

            // Filtrar controle (gastos) para pessoa e mesAno
            var controlePessoa = controle
                .Where(c => c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && c.mesAno == mesAno.ToString())
                .ToList();

            // Agrupar gastos por cartão
            var gastosPorCartao = controlePessoa
                .GroupBy(c => string.IsNullOrEmpty(c.Cartao) ? "Outros" : c.Cartao)
                .Select(g => new
                {
                    Cartao = g.Key,
                    Valor = g.Sum(x => x.Valor)
                }).ToList();

            var totalGastoControle = controlePessoa.Sum(c => c.Valor);

            var saldoFinal = (salario + extra) - fixosPessoa - totalGastoControle;


            var gastosPorCartaoDict = controlePessoa
                                    .GroupBy(c => string.IsNullOrEmpty(c.Cartao) ? "Outros" : c.Cartao)
                                    .ToDictionary(g => g.Key, g => g.Sum(x => x.Valor));


            var resumoPorCartaoTipo = controle
                                     .Where(c =>
                                     {
                                         var cartaoBase = RemoverPalavras(c.Cartao ?? "");

                                         // Se for Bradesco, só considera se a pessoa for igual à informada
                                         if (cartaoBase.Equals("Bradesco", StringComparison.OrdinalIgnoreCase))
                                         {
                                             return c.mesAno == mesAno &&
                                                    c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) &&
                                                    !string.IsNullOrEmpty(c.Cartao);
                                         }

                                         // Caso contrário, considera todas as pessoas
                                         return c.mesAno == mesAno && !string.IsNullOrEmpty(c.Cartao);
                                     })
                                     .GroupBy(c =>
                                     {
                                         var cartaoBase = RemoverPalavras(c.Cartao);

                                         var tipo = c.Cartao.Contains("Adicional", StringComparison.OrdinalIgnoreCase)
                                             ? "Adicional"
                                             : c.Cartao.Contains("Titular", StringComparison.OrdinalIgnoreCase)
                                                 ? "Titular"
                                                 : "Outro";

                                         return new { Cartao = cartaoBase, Tipo = tipo };
                                     })
                                     .Select(g => new CartaoTipoResumo
                                     {
                                         Cartao = g.Key.Cartao,
                                         Tipo = g.Key.Tipo,
                                         Valor = g.Sum(x => x.Valor)
                                     })
                                     .ToList();

            // Montar objeto retorno
            var resultado = new
            {
                Pessoa = pessoa,
                Periodo = mesAno,
                Salario = salario,
                Extras = extra,
                GastosFixos = fixosPessoa,
                TotalGasto = controlePessoa.Sum(c => c.Valor),
                SaldoRestante = (salario + extra) - fixosPessoa - controlePessoa.Sum(c => c.Valor),
                // SaldoCritico será calculado no front, pode deixar false aqui
                SaldoCritico = false,
                Compras = controlePessoa.Select(c => new Dictionary<string, object>
                {
                    ["IdLan"] = c.IdLan,
                    ["FormaPgto"] = c.FormaPgto,
                    ["Parcela"] = c.Parcela,
                    ["Compra"] = c.Compra,
                    ["Valor"] = c.Valor,
                    ["MesAno"] = c.mesAno,
                    ["Data"] = c.Data,
                    ["Pessoa"] = c.Pessoa,
                    ["Fonte"] = c.Fonte,
                    ["Cartao"] = c.Cartao
                }).ToList(),
                ResumoPorCartao = gastosPorCartaoDict,
                ResumoPorCartaoTipo = resumoPorCartaoTipo
            };

            return (true, "Sucesso", resultado);
        }
        catch (Exception ex)
        {
            return (false, ex.Message, null);
        }
    }
    public async Task DeletarLinhaPorIdAsync(string id, string nomeBase)
    {
        try
        {
            // Passo 1: Obtemos todas as linhas da aba
            var range = $"{nomeBase}!A:H";
            var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = await request.ExecuteAsync();
            var valores = response.Values;

            if (valores == null || valores.Count == 0)
                return;

            // Passo 2: Identificamos as linhas que têm o id na primeira coluna (coluna A)
            var linhasParaExcluir = new List<int>();
            for (int i = 0; i < valores.Count; i++)
            {
                if (valores[i].Count > 0 && valores[i][0]?.ToString() == id)
                {
                    // Salvamos o índice real da linha na planilha (linha 1-based)
                    linhasParaExcluir.Add(i + 1);
                }
            }

            if (!linhasParaExcluir.Any())
                return;

            // Passo 3: Excluímos as linhas de baixo para cima
            linhasParaExcluir.Sort();
            linhasParaExcluir.Reverse();

            foreach (var linha in linhasParaExcluir)
            {
                var deleteRequest = new Request
                {
                    DeleteDimension = new DeleteDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = await ObterSheetIdPorNome(nomeBase),
                            Dimension = "ROWS",
                            StartIndex = linha - 1, // zero-based
                            EndIndex = linha       // exclusive
                        }
                    }
                };

                var batchRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request> { deleteRequest }
                };

                await _service.Spreadsheets.BatchUpdate(batchRequest, SpreadsheetId).ExecuteAsync();
            }
        }
        catch (Exception ex)
        {
            // Logue ou trate o erro conforme necessário
            Console.WriteLine($"Erro ao excluir linhas com ID '{id}': {ex.Message}");
            throw;
        }
    }

    private async Task<int> ObterSheetIdPorNome(string nomeAba)
    {
        var spreadsheet = await _service.Spreadsheets.Get(SpreadsheetId).ExecuteAsync();
        var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title.Equals(nomeAba, StringComparison.OrdinalIgnoreCase));

        if (sheet == null)
            throw new Exception($"A aba '{nomeAba}' não foi encontrada na planilha.");

        return (int)sheet.Properties.SheetId;
    }


    public int? ObterIndiceDaLinhaPorId(string idLan)
    {
        var linhas = ReadData($"{SheetName}!A:J"); // Ajuste o range conforme sua planilha

        if (linhas == null || linhas.Count <= 1)
            return null;

        var header = linhas[0];
        int indexIdLan = header
            .Select((valor, index) => new { valor, index })
            .FirstOrDefault(x => (x.valor?.ToString()?.Trim() ?? "").Equals("IdLan", StringComparison.OrdinalIgnoreCase))
            ?.index ?? -1;

        if (indexIdLan == -1)
            return null; // Coluna IdLan não encontrada

        for (int i = 1; i < linhas.Count; i++)
        {
            var linha = linhas[i];
            var idValor = linha.ElementAtOrDefault(indexIdLan)?.ToString()?.Trim();

            if (idValor == idLan)
            {
                return i + 1; // Soma 1 porque o índice na planilha começa em 1 (linha real)
            }
        }

        return null;
    }
    public bool EditarCompraNaPlanilha(string idLan, EditarCompraRequest request)
    {
        var indiceLinha = ObterIndiceDaLinhaPorId(idLan);

        if (indiceLinha == null)
            return false;

        var dadosAtualizados = new List<object?>
                                {
                                    request.IdLan,
                                    request.FormaPgto,
                                    request.Parcela,
                                    request.Compra,
                                    request.Valor,
                                    request.MesAno,
                                    request.Data?.ToString("dd/MM/yyyy"),
                                    request.Pessoa,
                                    request.Fonte == "" ? "Salario" : request.Fonte,
                                    request.Cartao

                                };

        var range = $"{SheetName}!A{indiceLinha}:J{indiceLinha}";

        AtualizaCompra(range, new List<IList<object?>> { dadosAtualizados });

        return true;
    }

    public void AtualizaCompra(string range, IList<IList<object>> values)
    {
        var request = _service.Spreadsheets.Values.Update(
            new Google.Apis.Sheets.v4.Data.ValueRange() { Values = values },
            SpreadsheetId,
            range);

        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        request.Execute();
    }

    public static string RemoverPalavras(string texto)
    {
        if (string.IsNullOrEmpty(texto))
            return texto;

        var palavrasParaRemover = new List<string> { "Titular", "Adicional" };

        foreach (var palavra in palavrasParaRemover)
        {
            texto = System.Text.RegularExpressions.Regex.Replace(
                texto,
                System.Text.RegularExpressions.Regex.Escape(palavra),
                "",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        return texto.Trim();
    }

    public decimal ParseDecimal(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return 0;

        input = input.Replace("R$", "").Trim();
        input = input.Replace(" ", "");

        // Trata o caso comum: número no formato BR (ex: "1.380,00")
        if (input.Contains(","))
        {
            // Remove separador de milhar (.) e troca vírgula decimal por ponto
            input = input.Replace(".", "").Replace(",", ".");
        }
        else
        {
            // Remove separador de milhar (,) se houver
            input = input.Replace(",", "");
        }

        if (decimal.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
            return valor;

        return 0;
    }

    public void WriteEntrada(List<object> entrada)
    {
        var valueRange = new ValueRange { Values = new List<IList<object>> { entrada } };

        var appendRequest = _service.Spreadsheets.Values.Append(
            valueRange, SpreadsheetId, "Config!A:F");

        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        appendRequest.Execute();
    }

    public bool PessoaTemEntradaCadastrada(string pessoa, DateTime dataCompra)
    {
        // Obter o mesAno baseado na dataCompra e na regra do dia 08
        var mesAno = ObterMesAnoCompetencia(dataCompra);

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

    /// <summary>
    /// Calcula o mes/ano da competência financeira, onde o mês vira no dia 08.
    /// </summary>
    public string ObterMesAnoCompetencia(DateTime data)
    {
        // Se o dia da data for menor que 8, volta para o mês anterior
        var mes = data.Day >= 4 || data.Day >= 8 ? data.Month : data.AddMonths(-1).Month;
        var ano = data.Day >= 4 || data.Day >= 8 ? data.Year : data.AddMonths(-1).Year;

        return $"{mes:D2}/{ano}";
    }

    // Busca entrada pelo nome da pessoa e mês
    public Dictionary<string, string>? GetEntradaPorPessoaEMes(string pessoa, string mesAno)
    {
        var range = "Config!A:F";
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        var values = response.Values;

        if (values == null || values.Count == 0)
            return null;

        // Supondo que a primeira linha é cabeçalho
        for (int i = 1; i < values.Count; i++)
        {
            var row = values[i];
            if (row.Count >= 6)
            {
                var pessoaCell = row[0]?.ToString();
                var mesAnoCell = row[3]?.ToString();

                if (string.Equals(pessoaCell, pessoa, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(mesAnoCell, mesAno, StringComparison.OrdinalIgnoreCase))
                {
                    return new Dictionary<string, string>
                    {
                        ["Pessoa"] = pessoaCell ?? "",
                        ["Fonte"] = row[1]?.ToString() ?? "",
                        ["Valor"] = row[2]?.ToString() ?? "0",
                        ["MesAno"] = mesAnoCell ?? "",
                        ["ValorHora"] = row[4]?.ToString() ?? "0",
                        ["Extras"] = row[5]?.ToString() ?? "0",
                        ["Linha"] = i.ToString()
                    };
                }
            }
        }

        return null;
    }


    // Atualiza a coluna Extras da entrada existente para pessoa e mês
    public void AtualizarExtrasEntrada(string pessoa, string mesAno, decimal novosExtras)
    {
        var entrada = GetEntradaPorPessoaEMes(pessoa, mesAno);
        if (entrada == null)
            throw new Exception("Entrada não encontrada para atualização.");

        int linha = int.Parse(entrada["Linha"]);
        int planilhaLinha = linha + 1; // Conta o cabeçalho da planilha

        string cell = $"F{planilhaLinha}"; // coluna F = Extras

        var valueRange = new ValueRange
        {
            Values = new List<IList<object>> { new List<object> { novosExtras } }
        };

        var updateRequest = _service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, $"Config!{cell}");
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.Execute();
    }

    public (DateTime inicio, DateTime fim) ObterPeriodoCortePorCartao(DateTime dataReferencia, string cartao)
    {
        int diaInicio = 8;
        int diaFim = 7;

        switch (cartao.Trim().ToUpperInvariant())
        {
            case "ITAÚ":
            case "ITAU":
                diaInicio = 8;
                diaFim = 7;
                break;
            case "BRADESCO":
                diaInicio = 5;
                diaFim = 4;
                break;
            case "SANTANDER":
                diaInicio = 11;
                diaFim = 10;
                break;
            default:
                diaInicio = 8;
                diaFim = 7;
                break;
        }

        // Define o mês base para o início da competência
        DateTime dataBase = dataReferencia.Day >= diaInicio
            ? new DateTime(dataReferencia.Year, dataReferencia.Month, diaInicio)
            : new DateTime(dataReferencia.AddMonths(-1).Year, dataReferencia.AddMonths(-1).Month, diaInicio);

        // fim é diaFim do mês seguinte ao mês base
        DateTime fim = new DateTime(dataBase.Year, dataBase.Month, 1).AddMonths(1); // primeiro dia do próximo mês
        fim = new DateTime(fim.Year, fim.Month, diaFim);

        return (dataBase, fim);
    }

    public async Task<List<string>> GetCartoesAsync()
    {
        var range = $"{CartoesSheet}!A2:A"; // Exemplo: começa da célula A2 (ignora cabeçalho)
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = await request.ExecuteAsync();
        var values = response.Values;

        var result = new List<string>();
        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                result.Add(row[0].ToString());
            }
        }

        return result;
    }

#region FIXOS

    public async Task DeletarLinha(int rowIndex)
    {
        var request = _service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, $"{"Fixos"}!A{rowIndex}:H{rowIndex}");
        await request.ExecuteAsync();
    }
    public async Task AdicionarLinhas(IList<IList<object>> valores)
    {
        var valueRange = new ValueRange
        {
            Values = valores
        };

        var appendRequest = _service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, rangeFixo);
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        await appendRequest.ExecuteAsync();
    }

    public async Task InserirLinhaAsync(LinhaGastoModel model)
    {
        var valores = new List<IList<object>>
                    {
                        new List<object>
                        {
                            model.Id,
                            model.Tipo,
                            model.MesAno,
                            model.Pessoa,
                            model.Vencimento,
                            model.Valor.ToString(),
                            model.Pago == false ? "Não" : "Sim",
                            model.Dividido
                        }
                    };

        await AdicionarLinhas(valores);
    }

    public async Task<IList<IList<object>>> ObterValores()
    {
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, rangeFixo);
        var response = await request.ExecuteAsync();
        return response.Values;
    }

    public async Task AtualizarLinha(int linhaNumero, IList<object> valores)
    {
        var range = $"Fixos!A{linhaNumero}:H{linhaNumero}"; // faixa da linha completa, ajuste colunas se quiser
        var valueRange = new ValueRange
        {
            Values = new List<IList<object>> { valores }
        };

        var updateRequest = _service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        await updateRequest.ExecuteAsync();
    }

    public (bool Success, string Message, LinhaGastoModel? Data) GetLinhaPorId(string id)
    {
        try
        {
            var linhas = ReadData(rangeFixo);

            if (linhas == null || linhas.Count <= 1)
                return (false, "Nenhum dado encontrado na planilha.", null);

            var header = linhas[0];

            for (int i = 1; i < linhas.Count; i++)
            {
                var linha = linhas[i];

                var idLinha = linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";

                if (idLinha.Equals(id, StringComparison.OrdinalIgnoreCase))
                {
                    var gasto = new LinhaGastoModel
                    {
                        Id = Convert.ToInt64(idLinha),
                        Tipo = linha.ElementAtOrDefault(1)?.ToString() ?? "",
                        MesAno = linha.ElementAtOrDefault(2)?.ToString() ?? "",
                        Pessoa = linha.ElementAtOrDefault(3)?.ToString() ?? "",
                        Vencimento = linha.ElementAtOrDefault(4)?.ToString(),
                        Valor = ParseDecimal(linha.ElementAtOrDefault(5)?.ToString()),
                        Pago = (linha.ElementAtOrDefault(6)?.ToString()?.Trim().ToLower() == "true"),
                        Dividido = linha.ElementAtOrDefault(7)?.ToString()
                    };

                    return (true, "Linha encontrada com sucesso.", gasto);
                }
            }

            return (false, "Linha não encontrada.", null);
        }
        catch (Exception ex)
        {
            return (false, $"Erro ao acessar a planilha: {ex.Message}", null);
        }
    }

    public async Task<bool> AtualizarLinhaPorIdAsync(string id, decimal novoValor, string Dividido)
    {
        // Lê todas as linhas da aba
        var linhas = ReadData(rangeFixo);

        if (linhas == null || linhas.Count <= 1)
            return false;

        var header = linhas[0]; // Cabeçalho

        // Procura a linha pelo ID (assumindo que ID está na coluna 0)
        for (int i = 1; i < linhas.Count; i++)
        {
            var linha = linhas[i];
            var idLinha = linha.ElementAtOrDefault(0)?.ToString()?.Trim();

            if (string.Equals(idLinha, id, StringComparison.OrdinalIgnoreCase))
            {
                int numeroDaLinhaNaPlanilha = i + 1; // +1 porque a contagem na planilha começa em 1

                // Garante que a linha tenha pelo menos 7 colunas
                while (linha.Count < 8)
                    linha.Add("");

                // Atualiza o valor na coluna 5 (índice 5, que é a 6ª coluna → "Valor")
                linha[5] = novoValor.ToString(CultureInfo.InvariantCulture).ToString();
                linha[7] = Dividido;

                await AtualizarLinha(numeroDaLinhaNaPlanilha, linha);
                return true;
            }
        }

        return false; // Não encontrou o ID
    }

    public async Task<List<string>> GetFixosTipoAsync()
    {
        var range = $"{FixosTipoSheet}!A2:A"; // Exemplo: começa da célula A2 (ignora cabeçalho)
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = await request.ExecuteAsync();
        var values = response.Values;

        var result = new List<string>();
        if (values != null && values.Count > 0)
        {
            foreach (var row in values)
            {
                result.Add(row[0].ToString());
            }
        }

        return result;
    }

    #endregion


    internal class CartaoTipoResumo
    {
        public string? Cartao { get; set; }
        public string? Tipo { get; set; }
        public decimal Valor { get; set; }
    }
    public class ResumoPessoaDTO
    {
        public string Pessoa { get; set; } = "";
        public decimal SaldoRestante { get; set; }
        public string UltimaCompra { get; set; } = "-";
        public string DescricaoUltimaCompra { get; set; } = "-";
    }
}
