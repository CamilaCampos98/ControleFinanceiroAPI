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

                throw new Exception(ex.Message);
            }
            
        }
    }
}
