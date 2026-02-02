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


    public (bool Success, string Message, List<ResumoPessoaMesDTO>? Data) ResumoGeralPorMes()
    {
        try
        {
            var linhasControle = ReadData("Controle!A1:J");
            var configData = ReadData("Config!A1:F");
            var fixosData = ReadData("Fixos!A1:H");

            if (linhasControle == null || configData == null || fixosData == null)
                return (false, "Dados insuficientes nas planilhas.", null);

            var hoje = DateTime.Today;
            DateTime mesInicio = hoje.Day <= 8 ? new DateTime(hoje.Year, hoje.Month, 1).AddMonths(-1) : new DateTime(hoje.Year, hoje.Month, 1);
            int anoAtual = mesInicio.Year;

            //10 meses a partir de hoje
            var meses = Enumerable.Range(0, 6)
                       .Select(i => mesInicio.AddMonths(i))
                       .ToList();

            var pessoas = configData.Skip(1)
                .Select(l => l.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "")
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var resultado = new List<ResumoPessoaMesDTO>();

            foreach (var pessoa in pessoas)
            {
                foreach (var dataMes in meses)
                {
                    string mesAnoStr = dataMes.ToString("MM/yyyy");

                    var resumo = GetResumoPorPessoaEPeriodoInternal(pessoa, mesAnoStr, (List<IList<object>>)configData, (List<IList<object>>)fixosData, (List<IList<object>>)linhasControle);
                    if (resumo.Success && resumo.Data != null)
                    {
                        resultado.Add(new ResumoPessoaMesDTO
                        {
                            Pessoa = pessoa,
                            MesAno = mesAnoStr,
                            SaldoRestante = resumo.Data.SaldoRestante,
                            ValorGuardado = resumo.Data.ValorGuardado 
                        });
                    }
                }
            }

            return (true, "Sucesso", resultado.OrderBy(x => x.Pessoa).ThenBy(x => x.MesAno).ToList());
        }
        catch (Exception ex)
        {
            return (false, $"Erro: {ex.Message}", null);
        }
    }

    private (bool Success, string Message, dynamic Data) GetResumoPorPessoaEPeriodoInternal(
    string pessoa,
    string mesAno,
    List<IList<object>> configData,
    List<IList<object>> fixosData,
    List<IList<object>> controleData)
    {
        try
        {
            var configs = configData.Skip(1).Select(row => new
            {
                Pessoa = row[0]?.ToString() ?? "",
                Fonte = row[1]?.ToString() ?? "",
                Valor = ParseDecimal(row[2]?.ToString()),
                mesAno = row[3]?.ToString() ?? "",
                ValorHora = ParseDecimal(row[4]?.ToString()),
                Extras = ParseDecimal(row[5]?.ToString())
            }).ToList();

            var fixos = fixosData.Skip(1).Select(row => new
            {
                Id = row[0]?.ToString(),
                Tipo = row[1]?.ToString(),
                mesAno = row[2]?.ToString(),
                Pessoa = row[3]?.ToString(),
                Vencimento = row[4]?.ToString(),
                Valor = ParseDecimal(row[5]?.ToString()),
                Pago = row[6]?.ToString(),
                Dividido = row[7]?.ToString()
            }).ToList();

            var controle = controleData.Skip(1).Select(row => new
            {
                IdLan = row[0]?.ToString(),
                FormaPgto = row[1]?.ToString(),
                Parcela = row[2]?.ToString(),
                Compra = row[3]?.ToString(),
                Valor = ParseDecimal(row[4]?.ToString()),
                mesAno = row[5]?.ToString(),
                Data = DateTime.TryParse(row[6]?.ToString(), out var dt) ? dt : DateTime.MinValue,
                Pessoa = row[7]?.ToString(),
                Fonte = row[8]?.ToString(),
                Cartao = row[9]?.ToString()
            }).ToList();

            // Fechamento fixo por cartão
            var fechamentoCartoes = new Dictionary<string, int>
        {
            { "ITAU", 8 },
            { "SANTANDER", 9 },
            { "BRADESCO", 3 },
            { "C&A", 5 },
            { "RIACHUELO", 10 }
        };

            // Mes/Ano
            if (!DateTime.TryParseExact(mesAno, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoDate))
                throw new Exception("mesAno inválido. Formato esperado: MM/yyyy (ex: 05/2025).");

            // Função para calcular período da fatura
            (DateTime inicio, DateTime fim) ObterPeriodoCartao(string cartaoBase)
            {
                cartaoBase = RemoverPalavras(cartaoBase?.ToUpper() ?? "");
                if (!fechamentoCartoes.ContainsKey(cartaoBase))
                    cartaoBase = "OUTROS"; // default
                int diaFechamento = fechamentoCartoes.ContainsKey(cartaoBase) ? fechamentoCartoes[cartaoBase] : 1;

                // Início: dia seguinte ao fechamento no mês atual
                int diasNoMesAtual = DateTime.DaysInMonth(mesAnoDate.Year, mesAnoDate.Month);
                int diaInicio = Math.Min(diaFechamento, diasNoMesAtual);
                DateTime inicio = new DateTime(mesAnoDate.Year, mesAnoDate.Month, diaInicio).AddDays(1);

                // Fim: dia do fechamento no mês seguinte
                DateTime mesSeguinte = mesAnoDate.AddMonths(1);
                int diasNoMesSeguinte = DateTime.DaysInMonth(mesSeguinte.Year, mesSeguinte.Month);
                int diaFim = Math.Min(diaFechamento, diasNoMesSeguinte);
                DateTime fim = new DateTime(mesSeguinte.Year, mesSeguinte.Month, diaFim);

                return (inicio, fim);
            }

            // Salário e extras
            var dadosConfig = configs
                .Where(c => c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && c.mesAno == mesAno)
                .ToList();

            decimal salario = dadosConfig.Where(c => c.Fonte.Equals("Salario", StringComparison.OrdinalIgnoreCase)).Sum(c => c.Valor);
            decimal extras = dadosConfig.Sum(c => c.Extras);

            // Fixos
            decimal fixosPessoa = fixos
                .Where(f => f.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && f.mesAno == mesAno)
                .Sum(f => f.Valor);

            // Compras da pessoa no período
            var controlePessoa = controle
                .Where(c =>
                {
                    if (!c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase))
                        return false;

                    var (inicio, fim) = ObterPeriodoCartao(c.Cartao);
                    return c.Data.Date >= inicio && c.Data.Date <= fim;
                })
                .ToList();

            decimal totalGastoControle = controlePessoa.Sum(c => c.Valor);

            // Guardado
            decimal valorGuardado = fixos
                .Where(f => f.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) &&
                            f.mesAno == mesAno &&
                            f.Tipo?.IndexOf("guardado", StringComparison.OrdinalIgnoreCase) >= 0)
                .Sum(f => f.Valor);

            var fixosSemGuardado = fixos
                .Where(f => f.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) &&
                            f.mesAno == mesAno &&
                            f.Tipo?.IndexOf("guardado", StringComparison.OrdinalIgnoreCase) < 0)
                .Sum(f => f.Valor);

            decimal saldoFinal = (salario + extras) - fixosPessoa - totalGastoControle;

            return (true, "Sucesso", new
            {
                SaldoRestante = saldoFinal,
                ValorGuardado = valorGuardado
            });
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
            if (string.IsNullOrEmpty(pessoa))
                throw new Exception("Pessoa é obrigatória.");

            if (!DateTime.TryParseExact(mesAno, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoDate))
                throw new Exception("mesAno inválido. Formato esperado: MM/yyyy (ex: 05/2025).");

            // Ler dados das planilhas
            var configData = ReadData("Config!A1:F");
            var fixosData = ReadData("Fixos!A1:H");
            var controleData = ReadData("Controle!A1:J");

            // Mapear Config
            var configs = configData.Skip(1).Select(row => new
            {
                Pessoa = row[0].ToString(),
                Fonte = row[1].ToString(),
                Valor = ParseDecimal(row[2].ToString()),
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

            // Fechamento fixo por cartão
            var fechamentoCartoes = new Dictionary<string, int>
        {
            { "ITAU", 8 },
            { "SANTANDER", 9 },
            { "BRADESCO", 3 },
            { "C&A", 5 },
            { "RIACHUELO", 10 },
            { "OUTROS", 1 }
        };

            // Função para calcular período da fatura
            (DateTime inicio, DateTime fim) ObterPeriodoCartao(string cartaoBase)
            {
                cartaoBase = RemoverPalavras(cartaoBase?.ToUpper() ?? "");
                if (!fechamentoCartoes.ContainsKey(cartaoBase))
                    cartaoBase = "OUTROS";

                int diaFechamento = fechamentoCartoes[cartaoBase];

                // Início: dia seguinte ao fechamento no mês informado
                int diasMesAtual = DateTime.DaysInMonth(mesAnoDate.Year, mesAnoDate.Month);
                int diaInicio = Math.Min(diaFechamento, diasMesAtual);
                DateTime inicio = new DateTime(mesAnoDate.Year, mesAnoDate.Month, diaInicio).AddDays(1);

                // Fim: dia do fechamento no mês seguinte
                DateTime mesSeguinte = mesAnoDate.AddMonths(1);
                int diasMesSeguinte = DateTime.DaysInMonth(mesSeguinte.Year, mesSeguinte.Month);
                int diaFim = Math.Min(diaFechamento, diasMesSeguinte);
                DateTime fim = new DateTime(mesSeguinte.Year, mesSeguinte.Month, diaFim);

                return (inicio, fim);
            }

            // ----------------------------
            // FILTRAGENS
            // ----------------------------

            // Config da pessoa
            var dadosConfig = configs
                .Where(c => c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && c.mesAno == mesAno)
                .ToList();

            var salario = dadosConfig.Where(c => c.Fonte.Equals("Salario", StringComparison.OrdinalIgnoreCase)).Sum(c => c.Valor);
            var extra = dadosConfig.Sum(c => c.Extras);

            // Fixos da pessoa
            var fixosPessoa = fixos
                .Where(f => f.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) && f.mesAno == mesAno)
                .Sum(f => f.Valor);

            var controlePessoa = controle
                                        .Where(c =>
                                        {
                                            if (!c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase)) return false;

                                            var (inicio, fim) = ObterPeriodoCartao(c.Cartao);
                                            return c.Data.Date >= inicio && c.Data.Date <= fim;
                                        })
                                        .ToList();

            var gastosPorCartaoDict = controlePessoa
                .GroupBy(c => string.IsNullOrEmpty(c.Cartao) ? "Outros" : c.Cartao)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Valor));

            var resumoPorCartaoTipo = controle
    .Where(c =>
    {
        var cartaoBase = RemoverPalavras(c.Cartao ?? "").ToUpper();
        var (inicio, fim) = ObterPeriodoCartao(cartaoBase);

        // Se for BRADESCO, C&A ou RIACHUELO → filtra por pessoa
        if (cartaoBase == "BRADESCO" || cartaoBase == "C&A" || cartaoBase == "RIACHUELO")
        {
            return c.Pessoa.Equals(pessoa, StringComparison.OrdinalIgnoreCase) &&
                   c.Data.Date >= inicio && c.Data.Date <= fim &&
                   !string.IsNullOrEmpty(c.Cartao);
        }

        // Senão, traz todos para separar Titular e Adicional
        return c.Data.Date >= inicio && c.Data.Date <= fim &&
               !string.IsNullOrEmpty(c.Cartao);
    })
    .GroupBy(c =>
    {
        var cartaoBase = RemoverPalavras(c.Cartao).ToUpper();
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

            // ordem preferida dos cartões
            var preferredOrder = new[] { "ITAU", "SANTANDER", "RIACHUELO", "C&A", "BRADESCO" };

            // índice rápido para lookup
            var preferredIndex = preferredOrder
                .Select((name, idx) => new { Key = RemoverPalavras(name).ToUpper(), Index = idx })
                .ToDictionary(x => x.Key, x => x.Index);

            // aplica a ordenação
            resumoPorCartaoTipo = resumoPorCartaoTipo
                .OrderBy(r =>
                {
                    var key = RemoverPalavras(r.Cartao).ToUpper();
                    return preferredIndex.TryGetValue(key, out var idx) ? idx : int.MaxValue;
                })
                .ThenBy(r => r.Tipo == "Adicional" ? 0 : r.Tipo == "Titular" ? 1 : 2) // Adicional antes de Titular, depois Outro
                .ToList();


            // ----------------------------
            // RESULTADO FINAL
            // ----------------------------

            var totalGastoPessoa = controlePessoa.Sum(c => c.Valor);
            var saldoFinal = (salario + extra) - fixosPessoa - totalGastoPessoa;

            var resultado = new
            {
                Pessoa = pessoa,
                Periodo = mesAno,
                Salario = salario,
                Extras = extra,
                GastosFixos = fixosPessoa,
                TotalGasto = totalGastoPessoa,
                SaldoRestante = saldoFinal,
                SaldoCritico = saldoFinal < 0,
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
                ResumoPorCartao = gastosPorCartaoDict,          // Total do cartão (geral)
                ResumoPorCartaoTipo = resumoPorCartaoTipo
                                    .ToList()
        };

            return (true, "Sucesso", resultado);
        }
        catch (Exception ex)
        {
            return (false, ex.Message, null);
        }
    }



    public class CicloFatura
    {
        public string Cartao { get; set; }      // Nome do cartão, ex: ITAU
        public int Mes { get; set; }            // Mês do ciclo
        public int Ano { get; set; }            // Ano do ciclo
        public DateTime Fechamento { get; set; } // Data de fechamento da fatura
        public DateTime Vencimento { get; set; } // Data de vencimento da fatura
    }

    Dictionary<string, int> fechamentoCartoes = new()
                                                {
                                                    { "ITAU", 8 },
                                                    { "SANTANDER", 9 },
                                                    { "BRADESCO", 3 },
                                                    { "C&A", 5 },
                                                    { "RIACHUELO", 10 }
                                                };

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

    public async Task<List<UsuarioModel>> ObterUsuariosAsync()
    {
        var range = "usuarios!A2:B"; // Assume primeira linha como header
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = await request.ExecuteAsync();

        var result = new List<UsuarioModel>();
        foreach (var row in response.Values)
        {
            if (row.Count >= 2)
                result.Add(new UsuarioModel
                {
                    Usuario = row[0].ToString(),
                    Senha = row[1].ToString()
                });
        }

        return result;
    }

    public async Task<HashSet<string>> GetComprasAsync(string pessoaFiltro)
    {
        var request = _service.Spreadsheets.Values.Get(
            SpreadsheetId,
            $"{SheetName}!A:J");

        var response = await request.ExecuteAsync();
        var values = response.Values;

        var result = new HashSet<string>();

        if (values == null || values.Count <= 1)
            return result;

        var pessoaFiltroNorm = pessoaFiltro.Trim().ToUpperInvariant();

        // pula cabeçalho
        for (int i = 1; i < values.Count; i++)
        {
            var row = values[i];

            // Colunas conforme sua planilha:
            // E = Valor  -> index 4
            // G = Data   -> index 6
            // H = Pessoa -> index 7

            var valorStr = row.Count > 4 ? row[4]?.ToString() : null;
            var dataStr = row.Count > 6 ? row[6]?.ToString() : null;
            var pessoa = row.Count > 7 ? row[7]?.ToString() : null;

            if (string.IsNullOrWhiteSpace(valorStr) ||
                string.IsNullOrWhiteSpace(dataStr) ||
                string.IsNullOrWhiteSpace(pessoa))
                continue;

            // 👉 filtra pela pessoa já aqui
            if (!pessoa.Trim().ToUpperInvariant().Equals(pessoaFiltroNorm))
                continue;

            if (!DateTime.TryParse(dataStr, out var data))
                continue;

            if (!decimal.TryParse(
                    valorStr,
                    System.Globalization.NumberStyles.Any,
                    new System.Globalization.CultureInfo("pt-BR"),
                    out var valor))
                continue;

            var chave =
                $"{pessoa.Trim().ToUpperInvariant()}|" +
                $"{data:yyyy-MM-dd}|" +
                $"{valor.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

            result.Add(chave);
        }

        return result;
    }

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
