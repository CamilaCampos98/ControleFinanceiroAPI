using ControleFinanceiroAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ControleFinanceiroAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuariosController : Controller
    {
        private readonly GoogleSheetsService _googleSheetsService;

        public UsuariosController(GoogleSheetsService googleSheetsService)
        {
            _googleSheetsService = googleSheetsService;
        }

        [HttpGet("Login")]
        public async Task<IActionResult> ObterUsuarios()
        {
            try
            {
                var usuarios = await _googleSheetsService.ObterUsuariosAsync();
                return Ok(usuarios.Select(u => new { u.Usuario, u.Senha }).ToList());
            }
            catch (Exception ex)
            {
                // log...
                return StatusCode(500, $"Erro interno: {ex.Message}");
            }
        }


    }
}
