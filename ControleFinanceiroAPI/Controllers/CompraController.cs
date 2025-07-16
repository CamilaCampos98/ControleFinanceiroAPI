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
            var (success, message, data) = _googleSheetsService.ResumoGeralPorMes();

            if (!success)
                return BadRequest(message);

            return Ok(data);
        }

        [HttpGet("ResumoPessoaPeriodo")]
        public IActionResult GetResumoPorPessoaEPeriodo(string pessoa, string mesAno)
        {
            if (string.IsNullOrEmpty(pessoa))
                return BadRequest("Pessoa é obrigatória.");

            if (!DateTime.TryParseExact(mesAno, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime mesAnoDate))
                return BadRequest("mesAno inválido. Formato esperado: MM/yyyy (ex: 05/2025).");

            var (success, message, data) = _googleSheetsService.GetResumoPorPessoaEPeriodo(pessoa, mesAno);

            if (!success)
                return BadRequest(message);

            return Ok(data);
        }

        [HttpPost("DeletarPorId")]
        public async Task<IActionResult> DeletarPorId([FromBody] DeletarPorIdModel payload)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(payload.Id) || string.IsNullOrWhiteSpace(payload.NomeBase))
                {
                    return BadRequest(new { status = "erro", message = "Parâmetros obrigatórios não informados." });
                }

                await _googleSheetsService.DeletarLinhaPorIdAsync(payload.Id, payload.NomeBase);

                return Ok(new { status = "sucesso", message = $"Lançamentos com ID '{payload.Id}' deletados com sucesso." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
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
                var temEntrada = _googleSheetsService.PessoaTemEntradaCadastrada(compra.Pessoa, compra.Data);

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

                    decimal novosExtras = valorExtraCalculado;

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

        [HttpPut("EditarCompra")]
        public IActionResult Editar([FromBody] EditarCompraRequest request)
        {
            var sucesso = _googleSheetsService.EditarCompraNaPlanilha(request.IdLan, request);

            if (!sucesso)
                return NotFound("Compra não encontrada.");

            return Ok("Compra atualizada com sucesso.");
        }

        [HttpGet("GetCartoes")]
        public async Task<ActionResult<List<string>>> GetCartoes()
        {
            try
            {
                var cartoes = await _googleSheetsService.GetCartoesAsync();

                if (cartoes == null || cartoes.Count == 0)
                    return NoContent(); // 204

                return Ok(cartoes); // 200 com lista de strings
            }
            catch (Exception ex)
            {
                // Em produção, logue o erro
                return StatusCode(500, $"Erro ao buscar cartões: {ex.Message}");
            }
        }
        #region FIXOS
        [HttpGet("ListarFixos")]
        public async Task<IActionResult> ListarFixos([FromQuery] string pessoa, [FromQuery] string periodo)
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
                    if (linha[2]?.ToString() != periodo.ToString()) continue;

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

                var linhasAtuais = await _googleSheetsService.ObterValores();

                var linhasParaInserir = new List<IList<object>>();
                var fixosInseridos = new List<object>();
                var fixosIgnorados = new List<object>();

                foreach (var fixo in payload.Data)
                {
                    var jaExiste = linhasAtuais.Skip(1).Any(l =>
                    {
                        var tipo = l.ElementAtOrDefault(1)?.ToString()?.Trim();
                        var mesAno = l.ElementAtOrDefault(2)?.ToString()?.Trim();
                        var pessoa = l.ElementAtOrDefault(3)?.ToString()?.Trim();

                        return string.Equals(tipo, fixo.Tipo, StringComparison.OrdinalIgnoreCase) &&
                               string.Equals(mesAno, fixo.MesAno, StringComparison.OrdinalIgnoreCase) &&
                               string.Equals(pessoa, payload.Pessoa, StringComparison.OrdinalIgnoreCase);
                    });

                    if (jaExiste)
                    {
                        fixosIgnorados.Add(new { fixo.Tipo, fixo.MesAno, Pessoa = payload.Pessoa });
                        continue;
                    }

                    var linha = new List<object>
                                {
                                    fixo.Id,
                                    fixo.Tipo,
                                    fixo.MesAno,
                                    payload.Pessoa,
                                    fixo.Vencimento,
                                    fixo.Valor ?? "",
                                    fixo.Pago ?? "",
                                    fixo.Dividido
                                };

                    linhasParaInserir.Add(linha);
                    fixosInseridos.Add(new { fixo.Tipo, fixo.MesAno, Pessoa = payload.Pessoa });
                }

                if (linhasParaInserir.Any())
                {
                    await _googleSheetsService.AdicionarLinhas(linhasParaInserir);
                }

                return Ok(new
                {
                    status = "sucesso",
                    message = "Processo concluído com sucesso.",
                    inseridos = fixosInseridos,
                    ignorados = fixosIgnorados
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }


        [HttpPost("CopiarFixos")]
        public async Task<IActionResult> CopiarFixos([FromBody] CopiaFixosPayload payload)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(payload.Pessoa))
                    return BadRequest(new { status = "erro", message = "Pessoa não informada." });

                if (string.IsNullOrWhiteSpace(payload.MesAnoDestino))
                    return BadRequest(new { status = "erro", message = "Mês de destino não informado." });

                var linhas = await _googleSheetsService.ObterValores();

                var fixosAnteriores = linhas.Skip(1)
                    .Where(l =>
                        string.Equals(l.ElementAtOrDefault(3)?.ToString(), payload.Pessoa, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(l.ElementAtOrDefault(2)?.ToString(), payload.MesAnoOrigem, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (!fixosAnteriores.Any())
                    return NotFound(new { status = "erro", message = "Nenhum fixo encontrado no mês de origem." });

                // verifica se já existem para o mês destino
                var jaExistem = linhas.Skip(1)
                    .Any(l =>
                        string.Equals(l.ElementAtOrDefault(3)?.ToString(), payload.Pessoa, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(l.ElementAtOrDefault(2)?.ToString(), payload.MesAnoDestino, StringComparison.OrdinalIgnoreCase));

                if (jaExistem)
                    return Conflict(new { status = "erro", message = $"Já existem fixos cadastrados para {payload.MesAnoDestino}." });

                // Calcula vencimento baseado no mês de destino
                DateTime.TryParseExact(payload.MesAnoDestino, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dataDestino);
                var vencimento = new DateTime(dataDestino.Year, dataDestino.Month + 1, 10).ToString("yyyy-MM-dd");

                IList<IList<object>> novasLinhas = fixosAnteriores.Select(l => (IList<object>)new List<object>
                                                    {
                                                        DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + new Random().Next(1000, 9999),
                                                        l[1], // Tipo
                                                        payload.MesAnoDestino,
                                                        payload.Pessoa,
                                                        vencimento, // ✅ novo vencimento coerente com o mês destino
                                                        l[5], // Valor
                                                        l[6], // Pago
                                                        l[7]  // Dividido
                                                    }).ToList();

                await _googleSheetsService.AdicionarLinhas(novasLinhas);

                return Ok(new { status = "sucesso", message = $"Fixos copiados de {payload.MesAnoOrigem} para {payload.MesAnoDestino}." });
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
        public async Task<IActionResult> BuscaDataRef([FromQuery] string mesAno)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(mesAno))
                    return BadRequest(new { status = "erro", message = "O parâmetro 'mesAno' é obrigatório (formato MM/yyyy)." });

                // Validação de formato
                if (!DateTime.TryParseExact(mesAno, "MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dataRef))
                    return BadRequest(new { status = "erro", message = "Formato inválido para 'mesAno'. Use MM/yyyy." });

                var dados = _googleSheetsService.ReadData("Config!A:F");

                if (dados == null || dados.Count == 0)
                    return NotFound(new { status = "erro", message = "Nenhum dado encontrado na aba Config." });

                // Pega apenas a coluna da data (coluna D - index 3), ignora o cabeçalho
                var datas = dados
                    .Skip(1)
                    .Select(l => l.Count > 3 ? l[3]?.ToString()?.Trim() : null)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                // Verifica se o mês/ano informado já existe
                var existe = datas.Any(s => string.Equals(s, mesAno, StringComparison.OrdinalIgnoreCase));

                if (existe)
                    return Ok(new { status = "sucesso", mesAno = mesAno });
                else
                    return NotFound(new { status = "erro", message = $"O mês {mesAno} ainda não foi preenchido na aba Config." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { status = "erro", message = ex.Message });
            }
        }


        [HttpGet("GetTiposFixos")]
        public async Task<ActionResult<List<string>>> GetTipoFixos()
        {
            try
            {
                var fixos = await _googleSheetsService.GetFixosTipoAsync();

                if (fixos == null || fixos.Count == 0)
                    return NoContent(); // 204

                return Ok(fixos); // 200 com lista de strings
            }
            catch (Exception ex)
            {
                // Em produção, logue o erro
                return StatusCode(500, $"Erro ao buscar cartões: {ex.Message}");
            }
        }
        #endregion


    }
}

