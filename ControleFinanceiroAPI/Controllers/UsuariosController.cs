using ControleFinanceiroAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ControleFinanceiroAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuariosController : ControllerBase
    {
        private readonly GoogleSheetsService _googleSheetsService;
        private readonly ILogger<UsuariosController> _logger;

        public UsuariosController(GoogleSheetsService googleSheetsService, ILogger<UsuariosController> logger)
        {
            _googleSheetsService = googleSheetsService;
            _logger = logger;
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
                _logger.LogError(ex, "Erro ao obter usuários.");
                return StatusCode(500, "Erro interno ao obter usuários.");
            }
        }

        [HttpPost("Login")]
        public async Task<IActionResult> Login([FromBody] LoginRequest request)
        {
            if (request == null ||
                string.IsNullOrWhiteSpace(request.Usuario) ||
                string.IsNullOrWhiteSpace(request.Senha))
            {
                return BadRequest(new LoginResponse
                {
                    Autenticado = false,
                    Mensagem = "Usuário e senha são obrigatórios."
                });
            }

            try
            {
                var usuarios = await _googleSheetsService.ObterUsuariosAsync();
                var usuarioValido = usuarios.FirstOrDefault(u =>
                    string.Equals(u.Usuario, request.Usuario, StringComparison.OrdinalIgnoreCase) &&
                    u.Senha == request.Senha);

                if (usuarioValido == null)
                {
                    return Unauthorized(new LoginResponse
                    {
                        Autenticado = false,
                        Mensagem = "Login incorreto."
                    });
                }

                return Ok(new LoginResponse
                {
                    Autenticado = true,
                    Usuario = usuarioValido.Usuario,
                    Mensagem = "Login autorizado."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erro ao validar login do usuário {Usuario}.", request.Usuario);
                return StatusCode(500, new LoginResponse
                {
                    Autenticado = false,
                    Mensagem = "Erro interno ao validar login."
                });
            }
        }

    }
}
