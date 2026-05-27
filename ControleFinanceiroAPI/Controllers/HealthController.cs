using Microsoft.AspNetCore.Mvc;

namespace ControleFinanceiroAPI.Controllers;

[ApiController]
[Route("api/[controller]")]
public sealed class HealthController : ControllerBase
{
    [HttpGet]
    public IActionResult Get()
    {
        return Ok(new
        {
            status = "ok",
            service = "ControleFinanceiroAPI",
            timestamp = DateTimeOffset.UtcNow
        });
    }
}
