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
    private readonly string rangeFixo = "Fixos!A:G";

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
            var linhas = ReadData($"{SheetName}!A:J");

            if (linhas == null || linhas.Count <= 1)
                return (false, "Nenhum dado encontrado na planilha.", null);

            var header = linhas[0];

            // Leitura da aba Config para buscar o salário e extras
            var config = ReadData("Config!A:F");

            decimal salario = 0m;
            decimal extras = 0m;

            foreach (var linha in config.Skip(1)) // Ignora o header
            {
                var pessoaConfig = linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "";
                var mesAno = linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "";

                if (string.Equals(pessoaConfig, pessoa, StringComparison.OrdinalIgnoreCase) &&
                    (mesAno == $"{inicio:MM/yyyy}" || mesAno == $"{fim:MM/yyyy}"))
                {
                    var salarioStr = linha.ElementAtOrDefault(2)?.ToString();
                    salario = ParseDecimal(salarioStr);

                    var extrasStr = linha.ElementAtOrDefault(5)?.ToString(); // Coluna Extras, ajuste se precisar
                    extras = ParseDecimal(extrasStr);

                    break;
                }
            }

            if (salario == 0)
            {
                return (false, "Salário não encontrado na aba Config para essa pessoa e período.", null);
            }

            decimal totalRecebido = salario + extras;

            // Filtrar as compras
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

                if (dataCompra >= inicio && dataCompra <= fim)
                {
                    var valorStr = linha.ElementAtOrDefault(4)?.ToString();
                    var valor = ParseDecimal(valorStr);

                    // ----------- Preenche comprasFiltradas (filtrando por pessoa) ------------
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

                    // ----------- Agrupa por cartão base (sem titular/adicional) ------------
                    if (!string.IsNullOrEmpty(cartao))
                    {
                        string cartaoBase = RemoverPalavras(cartao); // Remove "Titular"/"Adicional"

                        // Se for cartão Bradesco, só soma se a pessoa for igual à do filtro
                        if (cartaoBase.Equals("Bradesco", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.Equals(pessoaLinha, pessoa, StringComparison.OrdinalIgnoreCase))
                            {
                                // Ignora essa compra para Bradesco se a pessoa for diferente
                                continue;
                            }
                        }

                        // Soma no total do cartão
                        if (resumoPorCartao.ContainsKey(cartaoBase))
                            resumoPorCartao[cartaoBase] += valor;
                        else
                            resumoPorCartao[cartaoBase] = valor;

                        // ----------- Agrupa também por tipo (Titular ou Adicional) ------------
                        string tipo = cartao.Contains("Adicional", StringComparison.OrdinalIgnoreCase) ? "Adicional" :
                                      cartao.Contains("Titular", StringComparison.OrdinalIgnoreCase) ? "Titular" : "Outro";

                        // Cria chave no formato CartaoBase|Tipo
                        string chave = $"{cartaoBase}|{tipo}";

                        if (resumoPorCartaoTipo.ContainsKey(chave))
                            resumoPorCartaoTipo[chave] += valor;
                        else
                            resumoPorCartaoTipo[chave] = valor;
                    }
                }
            }

            decimal totalGastoFixos = 0;
            // --- Novo: Ler dados da aba Fixos e somar os valores da coluna 5 para a mesma pessoa e período ---
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

                    // Verifica se o MesAno da linha fixa está dentro do período
                    // Como MesAno é MM/yyyy, vamos comparar com inicio e fim:
                    if (DateTime.TryParseExact(mesAnoFixo, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoData))
                    {
                        // Só soma se o MesAno estiver dentro do período (mes e ano)
                        if (mesAnoData >= new DateTime(inicio.Year, inicio.Month, 1) && mesAnoData <= new DateTime(fim.Year, fim.Month, 1))
                        {
                            var valorFixoStr = linhaFixo.ElementAtOrDefault(5)?.ToString(); // Coluna Valor índice 5
                            var valorFixo = ParseDecimal(valorFixoStr);
                            totalGasto += valorFixo;
                            totalGastoFixos += valorFixo;
                        }
                    }
                }
            }
            // --- Fim do novo ---

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

    public class TupleStringIgnoreCaseComparer : IEqualityComparer<(string, string)>
    {
        public bool Equals((string, string) x, (string, string) y)
        {
            return string.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.Item2, y.Item2, StringComparison.OrdinalIgnoreCase);
        }

        public int GetHashCode((string, string) obj)
        {
            return HashCode.Combine(
                obj.Item1?.ToLowerInvariant() ?? "",
                obj.Item2?.ToLowerInvariant() ?? ""
            );
        }
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

        // Remove espaços
        input = input.Replace(" ", "");

        // Conta quantas vírgulas e pontos existem
        int commaCount = input.Count(c => c == ',');
        int dotCount = input.Count(c => c == '.');

        // Cenário 1: ambos existem
        if (commaCount > 0 && dotCount > 0)
        {
            // Assume que o separador decimal é o último deles
            if (input.LastIndexOf(",") > input.LastIndexOf("."))
            {
                // Vírgula é decimal → ponto é milhar
                input = input.Replace(".", "");
                input = input.Replace(",", ".");
            }
            else
            {
                // Ponto é decimal → vírgula é milhar
                input = input.Replace(",", "");
            }
        }
        // Cenário 2: só vírgula
        else if (commaCount > 0)
        {
            if (input.LastIndexOf(",") >= input.Length - 3)
            {
                // vírgula é decimal
                input = input.Replace(".", "");
                input = input.Replace(",", ".");
            }
            else
            {
                // vírgula é milhar (raro, mas possível)
                input = input.Replace(",", "");
            }
        }
        // Cenário 3: só ponto
        else if (dotCount > 0)
        {
            if (input.LastIndexOf(".") >= input.Length - 3)
            {
                // ponto é decimal
                input = input.Replace(",", "");
            }
            else
            {
                // ponto é milhar
                input = input.Replace(".", "");
            }
        }

        if (decimal.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            return result;

        return 0;
    }

    internal class CartaoTipoResumo
    {
        public string? Cartao { get; set; }
        public string? Tipo { get; set; }
        public decimal Valor { get; set; }
    }
    public void WriteEntrada(List<object> entrada)
    {
        var valueRange = new ValueRange { Values = new List<IList<object>> { entrada } };

        var appendRequest = _service.Spreadsheets.Values.Append(
            valueRange, SpreadsheetId, "Config!A:F");

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

    #region FIXOS
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
        var range = $"Fixos!A{linhaNumero}:G{linhaNumero}"; // faixa da linha completa, ajuste colunas se quiser
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
                        Pago = (linha.ElementAtOrDefault(6)?.ToString()?.Trim().ToLower() == "true")
                        
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

    public async Task<bool> AtualizarLinhaPorIdAsync(string id, decimal novoValor)
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
                while (linha.Count < 7)
                    linha.Add("");

                // Atualiza o valor na coluna 5 (índice 5, que é a 6ª coluna → "Valor")
                linha[5] = novoValor.ToString(CultureInfo.InvariantCulture).ToString();

                await AtualizarLinha(numeroDaLinhaNaPlanilha, linha);
                return true;
            }
        }

        return false; // Não encontrou o ID
    }


    #endregion

}
