namespace ControleFinanceiroAPI.Models;

public sealed record PartyEvent(string Id, string Name, DateTime Date);
public sealed record PartyItem(string Id, string EventId, string Category, string Name, decimal? Amount);
public sealed record PartyInstallment(string Id, string EventId, string Category, string? ItemId, DateTime DueDate, decimal Amount);
public sealed record PartyPayment(string Id, string EventId, string Category, string? ItemId,
    string? InstallmentId, string Description, DateTime PaidAt, decimal Amount);

public sealed class PartySnapshot
{
    public PartyEvent? Event { get; init; }
    public List<PartyItem> Items { get; init; } = new();
    public List<PartyInstallment> Installments { get; init; } = new();
    public List<PartyPayment> Payments { get; init; } = new();
}

public sealed record PartyItemInput(string Category, string Name, decimal? Amount);
public sealed record PartyScheduleInput(string Category, string? ItemId, List<DateTime> Months,
    decimal MinimumInstallment, int DueDay);
