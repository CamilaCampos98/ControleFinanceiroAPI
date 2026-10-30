using ControleFinanceiroAPI.Models;
using Microsoft.AspNetCore.Mvc;

namespace ControleFinanceiroAPI.Controllers;

[ApiController]
[Route("api/Festas")]
public sealed class FestasController : ControllerBase
{
    private readonly GoogleSheetsService _sheets;

    public FestasController(GoogleSheetsService sheets) => _sheets = sheets;

    [HttpGet]
    public ActionResult<PartySnapshot> Get() => _sheets.ReadParty();

    [HttpPost("ImportarLivia")]
    public ActionResult<PartySnapshot> Import() => _sheets.ImportLiviaParty();

    [HttpPut("Lembrancinhas")]
    public ActionResult<PartySnapshot> UpdateSouvenirOrder([FromBody] PartySouvenirOrderInput input) =>
        Run(() => _sheets.UpdatePartySouvenirOrder(input));

    [HttpPost("Itens")]
    public ActionResult<PartySnapshot> AddItem([FromBody] PartyItemInput input) => Run(() => _sheets.AddPartyItem(input));

    [HttpPut("Itens/{id}")]
    public ActionResult<PartySnapshot> UpdateItem(string id, [FromBody] PartyItemInput input) =>
        Run(() => _sheets.UpdatePartyItem(id, input));

    [HttpPost("Parcelas")]
    public ActionResult<PartySnapshot> Schedule([FromBody] PartyScheduleInput input) =>
        Run(() => _sheets.SchedulePartyPayments(input));

    [HttpPost("Parcelas/{id}/Pagar")]
    public ActionResult<PartySnapshot> Pay(string id, [FromBody] PayInstallmentInput input) =>
        Run(() => _sheets.PayPartyInstallment(id, input.PaymentId, input.PaidAt, input.Amount));

    [HttpPost("Convidados")]
    public ActionResult<PartySnapshot> AddGuest([FromBody] PartyGuestInput input) =>
        Run(() => _sheets.AddPartyGuest(input));

    [HttpPut("Convidados/{id}")]
    public ActionResult<PartySnapshot> UpdateGuest(string id, [FromBody] PartyGuestInput input) =>
        Run(() => _sheets.UpdatePartyGuest(id, input));

    [HttpDelete("Convidados/{id}")]
    public ActionResult<PartySnapshot> DeleteGuest(string id) => Run(() => _sheets.DeletePartyGuest(id));

    [HttpPost("Casas")]
    public ActionResult<PartySnapshot> AddHouse([FromBody] PartyHouseInput input) =>
        Run(() => _sheets.AddPartyHouse(input));

    [HttpPut("Casas/{id}")]
    public ActionResult<PartySnapshot> UpdateHouse(string id, [FromBody] PartyHouseInput input) =>
        Run(() => _sheets.UpdatePartyHouse(id, input));

    [HttpDelete("Casas/{id}")]
    public ActionResult<PartySnapshot> DeleteHouse(string id) => Run(() => _sheets.DeletePartyHouse(id));

    private ActionResult<PartySnapshot> Run(Func<PartySnapshot> action)
    {
        try { return Ok(action()); }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        { return BadRequest(new { message = exception.Message }); }
    }

    public sealed record PayInstallmentInput(string PaymentId, DateTime PaidAt, decimal Amount);
}
