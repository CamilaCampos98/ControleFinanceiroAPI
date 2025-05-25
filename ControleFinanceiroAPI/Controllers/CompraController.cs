using ControleFinanceiroAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ControleFinanceiroAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ComprasController : ControllerBase
    {
        private readonly GoogleSheetsService _googleSheetsService;

        public ComprasController(GoogleSheetsService googleSheetsService)
        {
            _googleSheetsService = googleSheetsService;
        }

        [HttpPost]
        public IActionResult CadastrarCompra([FromBody] CompraModel compra)
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
    }
}
