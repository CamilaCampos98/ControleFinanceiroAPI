using ControleFinanceiroAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

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

                _googleSheetsService.WritePurchaseWithInstallments(compra);
                return Ok(new
                {
                    message = "Compra cadastrada com sucesso!",
                    compra
                });
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Erro no endpoint RegistrarCompra: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
            
        }

        [HttpGet("TodasCompras")]
        public IActionResult GetAllCompras()
        {
            try
            {
                // Lê todas as linhas da aba Controle (ajuste o range conforme sua necessidade)
                var linhas = _googleSheetsService.ReadData($"{SheetName}!A:E");

                if (linhas == null || linhas.Count == 0)
                    return NotFound("Nenhum dado encontrado na planilha.");

                return Ok(linhas);
            }
            catch (Exception ex)
            {
                // Log do erro se precisar
                return StatusCode(500, $"Erro ao acessar a planilha: {ex.Message}");
            }
        }

    }
}
