using ControleFinanceiroAPI.Models;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace ControleFinanceiroAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CompraController : ControllerBase
    {
        private readonly GoogleSheetsService _googleSheetsService;
        public string SheetName = "Controle";
        public CompraController(GoogleSheetsService googleSheetsService)
        {
            _googleSheetsService = googleSheetsService;
        }

        [HttpGet("Get")]
        public IActionResult Get()
        {
            return Ok("API Funcionando");
        }
        [HttpGet("ResumoGeral")]
        public IActionResult ResumoGeral()
        {
            var (success, message, data) = _googleSheetsService.ResumoGeral();

            if (!success)
                return BadRequest(message);

            return Ok(data);
        }

        [HttpGet("ResumoPessoaPeriodo")]
        public IActionResult GetResumoPorPessoaEPeriodo(string pessoa, string dataInicio, string dataFim)
        {
            if (!DateTime.TryParse(dataInicio, out DateTime inicio))
                return BadRequest("Data de início inválida.");

            if (!DateTime.TryParse(dataFim, out DateTime fim))
                return BadRequest("Data de fim inválida.");

            var (success, message, data) = _googleSheetsService.GetResumoPorPessoaEPeriodo(pessoa, inicio, fim);

            if (!success)
                return BadRequest(message);

            return Ok(data);
        }

        [HttpGet("TodasComprasPorPessoa")]
        public IActionResult GetAllComprasPorPessoa()
        {
            try
            {
                var linhas = _googleSheetsService.ReadData($"{SheetName}!A:I");

                if (linhas == null || linhas.Count <= 1)
                    return NotFound("Nenhum dado encontrado na planilha.");

                // A primeira linha é o cabeçalho
                var header = linhas[0];

                var comprasPorPessoa = new Dictionary<string, List<Dictionary<string, object>>>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i < linhas.Count; i++)
                {
                    var linha = linhas[i];

                    var pessoa = linha.ElementAtOrDefault(7)?.ToString() ?? "Desconhecido";

                    if (!comprasPorPessoa.ContainsKey(pessoa))
                        comprasPorPessoa[pessoa] = new List<Dictionary<string, object>>();

                    var compra = new Dictionary<string, object>();

                    for (int j = 0; j < header.Count; j++)
                    {
                        var chave = header[j]?.ToString() ?? $"Coluna{j}";
                        var valor = linha.ElementAtOrDefault(j) ?? "";
                        compra[chave] = valor;
                    }

                    comprasPorPessoa[pessoa].Add(compra);
                }

                return Ok(comprasPorPessoa);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Erro ao acessar a planilha: {ex.Message}");
            }
        }

        [HttpPost("RegistrarCompra")]
        public IActionResult CadastrarCompra([FromBody] CompraModel compra)
        {
            try
            {
                if (compra == null)
                    return BadRequest();

                var dataToWrite = new List<object>
                            {
                                compra.FormaPgto,
                                compra.TotalParcelas,
                                compra.Descricao,
                                compra.ValorTotal,
                                compra.Data.ToString("yyyy-MM-dd")
                            };

                // Verifica se a pessoa tem entrada no mês
                var temEntrada = _googleSheetsService.PessoaTemEntradaCadastrada(compra.Pessoa, compra.Data.ToString("MM/yyyy") ?? "");

                if (!temEntrada)
                {
                    return BadRequest($"A pessoa {compra.Pessoa} não possui entrada cadastrada no mês {compra.Data.ToString("MM/yyyy")}.");
                }

                _googleSheetsService.WritePurchaseWithInstallments(compra);
                return Ok(new { id = compra.idLan, message = "Compra registrada com sucesso" });

            }
            catch (Exception ex)
            {

                Console.WriteLine($"Erro no endpoint RegistrarCompra: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }

        }

        [HttpPost("RegistrarEntrada")]
        public IActionResult RegistrarEntrada([FromBody] EntradaModel entrada)
        {
            try
            {
                if (entrada == null)
                    return BadRequest("Dados inválidos");

                decimal valorCalculado = 0m;

                if (entrada.TipoEntrada != "Extra")
                {
                    // Cálculo padrão salário
                    valorCalculado = entrada.ValorHora * entrada.HorasUteisMes;

                    var linha = new List<object>
                            {
                                entrada.Pessoa,
                                entrada.TipoEntrada,
                                valorCalculado,
                                entrada.MesAno,
                                entrada.ValorHora,
                                entrada.HorasExtras,
                                "", // Extras fica vazio para salário
                            };

                    _googleSheetsService.WriteEntrada(linha);

                    return Ok(new
                    {
                        message = "Entrada registrada com sucesso!",
                        valorCalculado,
                        entrada
                    });
                }
                else if (entrada.TipoEntrada == "Extra")
                {
                    var entradaBase = _googleSheetsService.GetEntradaPorPessoaEMes(entrada.Pessoa, entrada.MesAno);

                    if (entradaBase == null)
                        return BadRequest("Entrada base (salário) não encontrada para a pessoa e mês.");

                    var valorString = entradaBase["ValorHora"]
                                    .Replace(".", "")
                                    .Replace(",", ".");

                    if (!decimal.TryParse(valorString, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal valorHoraExtra))
                        return BadRequest("Valor da hora inválido na base.");

                    // Calcular valor extra
                    decimal valorExtraCalculado = valorHoraExtra * entrada.HorasExtras;

                    // Pegar o valor extra atual para somar (coluna F)
                    var valorExtraString = entradaBase["Extras"]
                                    .Replace(".", "")
                                    .Replace(",", ".");

                    if (!decimal.TryParse(valorExtraString, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal extrasAtuais))
                        return BadRequest("Valor da extra inválido na base.");

                    decimal novosExtras = extrasAtuais + valorExtraCalculado;

                    // Atualizar a coluna Extras na planilha
                    _googleSheetsService.AtualizarExtrasEntrada(entrada.Pessoa, entrada.MesAno, novosExtras);

                    return Ok(new
                    {
                        message = "Horas extras registradas com sucesso!",
                        valorHoraExtra,
                        horasExtras = entrada.HorasExtras,
                        valorExtraCalculado,
                        novosExtras
                    });
                }
                else
                {
                    return BadRequest("Tipo de entrada inválido.");
                }


            }
            catch (Exception ex)
            {

                return StatusCode(500, $"Erro interno: {ex.Message}");
            }

        }

        #region FIXOS
        [HttpGet("ListarFixos")]
        public async Task<IActionResult> ListarFixos([FromQuery] string pessoa)
        {
            try
            {
                // Colunas da aba: Id, Tipo, MesAno, Pessoa, Vencimento, Valor, Pago
                var dados = await _googleSheetsService.ObterValores();

                if (dados == null || dados.Count == 0)
                    return Ok(new List<object>());

                var lista = new List<FixoModel>();

                // Pula o cabeçalho, começa do índice 1
                for (int i = 1; i < dados.Count; i++)
                {
                    var linha = dados[i];

                    // Garante que a linha tenha 7 colunas (preenche com null se faltar)
                    while (linha.Count < 7)
                        linha.Add(null);

                    // Se a coluna Pessoa (index 3) for diferente da pessoa buscada, pula
                    if (linha[3]?.ToString()?.Trim().ToUpper() != pessoa.ToUpper()) continue;

                    var fixo = new FixoModel
                    {
                        Id = long.TryParse(linha[0]?.ToString(), out var idVal) ? idVal : 0,
                        Tipo = linha[1]?.ToString(),
                        MesAno = linha[2]?.ToString(),
                        Vencimento = linha[4]?.ToString(),
                        Valor = linha[5]?.ToString()
                                    ?.Replace("R$", "", StringComparison.OrdinalIgnoreCase)
                                    ?.Replace(",", "")
                                    ?.Trim(),
                        Pago = linha[6]?.ToString()?.Trim(),
                        Dividido = linha[7]?.ToString()
                    };

                    lista.Add(fixo);
                }

                return Ok(lista);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }

        [HttpPost("GeraFixos")]
        public async Task<IActionResult> GeraFixos([FromBody] FixoPayload payload)
        {
            try
            {
                if (payload == null || payload.Data == null || !payload.Data.Any())
                    return BadRequest(new { status = "erro", message = "Payload inválido ou vazio." });

                if (string.IsNullOrEmpty(payload.Pessoa))
                    return BadRequest(new { status = "erro", message = "Campo 'Pessoa' é obrigatório." });

                // 🔍 Ler dados atuais da planilha
                var linhasAtuais = await _googleSheetsService.ObterValores();

                var linhasParaInserir = new List<IList<object>>();
                var fixosInseridos = new List<object>();
                var fixosIgnorados = new List<object>();

                foreach (var fixo in payload.Data)
                {
                    bool jaExiste = linhasAtuais.Skip(1).Any(l =>
                    {
                        var tipo = l.ElementAtOrDefault(1)?.ToString()?.Trim();
                        var mesAno = l.ElementAtOrDefault(2)?.ToString()?.Trim();
                        var pessoa = l.ElementAtOrDefault(3)?.ToString()?.Trim();

                        return string.Equals(tipo, fixo.Tipo, StringComparison.OrdinalIgnoreCase) &&
                               string.Equals(mesAno, fixo.MesAno, StringComparison.OrdinalIgnoreCase) &&
                               string.Equals(pessoa, payload.Pessoa, StringComparison.OrdinalIgnoreCase);
                    });

                    var resumo = new
                    {
                        fixo.Tipo,
                        fixo.MesAno,
                        Pessoa = payload.Pessoa
                    };

                    if (jaExiste)
                    {
                        fixosIgnorados.Add(resumo);
                        continue; // 🔸 Pula se já existe
                    }

                    var linha = new List<object>
                        {
                            fixo.Id,
                            fixo.Tipo,
                            fixo.MesAno,
                            payload.Pessoa,
                            fixo.Vencimento,
                            "",    // 🔸 Valor (a preencher depois)
                            "",    // 🔸 Pago (a preencher depois)
                            fixo.Dividido
                        };

                    linhasParaInserir.Add(linha);
                    fixosInseridos.Add(resumo);
                }

                if (linhasParaInserir.Any())
                {
                    await _googleSheetsService.AdicionarLinhas(linhasParaInserir);
                }

                return Ok(new
                {
                    status = "sucesso",
                    message = linhasParaInserir.Any() ? "Processo concluído com sucesso." : "Nenhum fixo novo para inserir.",
                    inseridos = fixosInseridos,
                    ignorados = fixosIgnorados
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }

        [HttpPost("DeletarFixo")]
        public async Task<IActionResult> DeletarFixo([FromBody] DeletarFixoPayload payload)
        {
            try
            {
                if (string.IsNullOrEmpty(payload.Id) ||
                    string.IsNullOrEmpty(payload.MesAno) ||
                    string.IsNullOrEmpty(payload.Pessoa))
                {
                    return BadRequest(new { status = "erro", message = "Campos obrigatórios não informados." });
                }

                var linhas = await _googleSheetsService.ObterValores();
                if (linhas == null || linhas.Count <= 1)
                    return NotFound(new { status = "erro", message = "Nenhuma linha encontrada." });

                // Localizar índice da linha que corresponde
                var index = linhas
                    .Select((linha, idx) => new { Linha = linha, Index = idx })
                    .FirstOrDefault(x =>
                        (x.Linha.ElementAtOrDefault(0)?.ToString() ?? "") == payload.Id &&
                        (x.Linha.ElementAtOrDefault(2)?.ToString() ?? "") == payload.MesAno &&
                        (x.Linha.ElementAtOrDefault(3)?.ToString() ?? "") == payload.Pessoa);

                if (index == null)
                    return NotFound(new { status = "erro", message = "Linha não encontrada." });

                // +1 porque na planilha o índice começa em 1 e pula o cabeçalho
                await _googleSheetsService.DeletarLinha(index.Index + 1);

                return Ok(new { status = "sucesso", message = "Fixo deletado com sucesso." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }

        [HttpPost("AtualizarFixo")]
        public async Task<IActionResult> AtualizarFixo([FromBody] AtualizaFixoModel model)
        {
            try
            {
                if (model == null || string.IsNullOrEmpty(model.Id))
                    return BadRequest(new { status = "erro", message = "Dados inválidos" });

                // Ler os dados atuais da planilha
                var dados = await _googleSheetsService.ObterValores();

                if (dados == null || dados.Count == 0)
                    return NotFound(new { status = "erro", message = "Planilha vazia" });

                // Encontrar a linha pelo Id (assumindo que Id está na coluna 0)
                var listaDados = dados.ToList();
                var linhaIndex = listaDados.FindIndex(row => row.Count > 0 && row[0]?.ToString() == model.Id);

                if (linhaIndex == -1)
                    return NotFound(new { status = "erro", message = "Fixo não encontrado" });

                var linha = dados[linhaIndex];

                // Garantir que linha tenha pelo menos 7 colunas
                while (linha.Count < 8)
                {
                    linha.Add(string.Empty);
                }
                // Atualizar valor e pago (colunas 5 e 6)
                linha[5] = model.Valor.ToString("F2").Replace(",", "."); // valor decimal formatado com ponto
                linha[6] = model.Pago ? "Sim" : "Não";
                linha[7] = model.Dividido ? "Sim" : "Não";
                
                // Atualiza a linha na planilha (assumindo que você tem método para atualizar uma linha específica)
                await _googleSheetsService.AtualizarLinha(linhaIndex + 1, linha);
                // +1 porque planilha geralmente indexa linha 1 como cabeçalho

                return Ok(new { status = "sucesso", message = "Fixo atualizado com sucesso" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }

        [HttpPost("DividirGasto")]
        public async Task<IActionResult> DividirGasto([FromBody] DividirGastoModel request)
        {
            var (success, message, linha) = _googleSheetsService.GetLinhaPorId(request.IdLinha);

            if (!success || linha == null)
                return NotFound(message);

            decimal valorAtual = linha.Valor;

            if (request.ValorDividir <= 0)
                return BadRequest("O valor deve ser maior que zero.");

            if (request.ValorDividir > valorAtual)
                return BadRequest("O valor para dividir é maior que o valor atual.");

            // 🔻 Atualiza o valor da linha original
            decimal novoValor = valorAtual - request.ValorDividir;
            await _googleSheetsService.AtualizarLinhaPorIdAsync(request.IdLinha, novoValor, request.Dividido);

            // ➕ Cria a nova linha para quem vai receber o valor dividido
            var novaLinha = new LinhaGastoModel
            {
                Id = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + new Random().Next(1000, 9999),
                Tipo = linha.Tipo,
                MesAno = linha.MesAno,
                Vencimento = linha.Vencimento = $"{DateTime.Today.Year}-{DateTime.Today.Month.ToString("D2")}-15",
                Valor = request.ValorDividir,
                Pago = linha.Pago,
                Pessoa = request.NomeDestino,
                Dividido = request.Dividido
            };

            await _googleSheetsService.InserirLinhaAsync(novaLinha);

            return Ok(new
            {
                mensagem = "Gasto dividido com sucesso.",
                linhaAtualizada = new { linha.Id, novoValor },
                novaLinha = novaLinha
            });
        }

        [HttpGet("BuscaDataRef")]
        public async Task<IActionResult> BuscaDataRef()
        {
            try
            {
                var dados = _googleSheetsService.ReadData("Config!A:F");

                if (dados == null || dados.Count == 0)
                    return NotFound(new { status = "erro", message = "Nenhum dado encontrado." });

                // Pega a coluna 3 (índice 2), ignora cabeçalhos e linhas sem valor
                var datas = dados
                    .Skip(1) // se a primeira linha é cabeçalho, senão remova essa linha
                    .Select(linha => linha.Count > 3 ? linha[3]?.ToString() : null)
                    .Where(mesAno => !string.IsNullOrWhiteSpace(mesAno))
                    .Select(mesAno =>
                    {
                        // Tenta converter para DateTime no formato MM/yyyy
                        if (DateTime.TryParseExact(mesAno, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                            return dt;
                        else
                            return (DateTime?)null;
                    })
                    .Where(dt => dt.HasValue)
                    .Select(dt => dt.Value)
                    .ToList();

                if (!datas.Any())
                    return NotFound(new { status = "erro", message = "Nenhuma data válida encontrada." });

                // Pega a data mais recente
                var dataMaisRecente = datas.Max();

                // Retorna no mesmo formato "MM/yyyy"
                var mesAnoMaisRecente = dataMaisRecente.ToString("MM/yyyy");

                return Ok(new { status = "sucesso", mesAno = mesAnoMaisRecente });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }
        #endregion


    }
}

