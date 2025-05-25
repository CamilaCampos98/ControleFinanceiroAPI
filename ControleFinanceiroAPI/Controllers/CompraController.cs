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

        [HttpGet]
        public IActionResult Get()
        {
            return Ok("API Funcionando");
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
    }
}
