namespace ControleFinanceiroAPI.Models;

public sealed record PartyEvent(string Id, string Name, DateTime Date, int? BoysOrdered = null, int? GirlsOrdered = null,
    int? AdultsOrdered = null);
public sealed record PartyItem(string Id, string EventId, string Category, string Name, decimal? Amount);
public sealed record PartyInstallment(string Id, string EventId, string Category, string? ItemId, DateTime DueDate, decimal Amount);
public sealed record PartyPayment(string Id, string EventId, string Category, string? ItemId,
    string? InstallmentId, string Description, DateTime PaidAt, decimal Amount);
public sealed record PartyGuest(string Id, string EventId, string Name, bool IsChild, string Gender,
    string Notes, string CreatedAt, string UpdatedAt, string HouseId);
public sealed record PartyHouse(string Id, string EventId, string Name, string CreatedAt, string UpdatedAt);

public sealed class PartySnapshot
{
    public PartyEvent? Event { get; init; }
    public List<PartyItem> Items { get; init; } = new();
    public List<PartyInstallment> Installments { get; init; } = new();
    public List<PartyPayment> Payments { get; init; } = new();
    public List<PartyGuest> Guests { get; init; } = new();
    public List<PartyHouse> Houses { get; init; } = new();
}

public sealed record PartyItemInput(string Category, string Name, decimal? Amount);
public sealed record PartyScheduleInput(string Category, string? ItemId, List<DateTime> Months,
    decimal MinimumInstallment, int DueDay);
public sealed record PartyGuestInput(string Name, bool IsChild, string? Gender, string? Notes, string? HouseId);
public sealed record PartyHouseInput(string Name, List<string>? GuestIds);
public sealed record PartySouvenirOrderInput(int BoysOrdered, int GirlsOrdered, int AdultsOrdered);
