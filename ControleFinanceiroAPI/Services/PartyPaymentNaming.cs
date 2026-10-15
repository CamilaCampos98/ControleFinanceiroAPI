using ControleFinanceiroAPI.Models;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ControleFinanceiroAPI.Services;

public static class PartyPaymentNaming
{
    private static readonly CultureInfo Brazil = CultureInfo.GetCultureInfo("pt-BR");

    public static string Describe(PartySnapshot snapshot, PartyInstallment installment, decimal amount,
        decimal previouslyPaid = 0)
    {
        var matchingInstallments = snapshot.Installments
            .Where(candidate => SameTarget(candidate.Category, candidate.ItemId,
                installment.Category, installment.ItemId))
            .OrderBy(candidate => candidate.DueDate).ThenBy(candidate => candidate.Id, StringComparer.Ordinal)
            .ToList();
        var position = matchingInstallments.FindIndex(candidate => candidate.Id == installment.Id);
        if (position < 0) throw new ArgumentException("Parcela não encontrada no planejamento.");

        var lastHistoricalNumber = snapshot.Payments
            .Where(payment => payment.InstallmentId == null &&
                SameTarget(payment.Category, payment.ItemId, installment.Category, installment.ItemId))
            .Select(payment => Regex.Match(payment.Description, @"^Parcela\s+(\d+)", RegexOptions.IgnoreCase))
            .Where(match => match.Success)
            .Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture))
            .DefaultIfEmpty(0).Max();
        var target = installment.ItemId == null ? installment.Category :
            snapshot.Items.FirstOrDefault(item => item.Id == installment.ItemId)?.Name ?? installment.Category;
        var description = $"Parcela {lastHistoricalNumber + position + 1} {target.ToLower(Brazil)}";
        return amount + previouslyPaid < installment.Amount ? description + " (pagamento parcial)" : description;
    }

    private static bool SameTarget(string leftCategory, string? leftItemId, string rightCategory, string? rightItemId) =>
        string.Equals(leftCategory, rightCategory, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(leftItemId, rightItemId, StringComparison.Ordinal);
}
