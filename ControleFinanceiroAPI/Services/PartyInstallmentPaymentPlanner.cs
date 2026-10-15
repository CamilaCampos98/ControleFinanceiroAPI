using ControleFinanceiroAPI.Models;

namespace ControleFinanceiroAPI.Services;

public static class PartyInstallmentPaymentPlanner
{
    public static IReadOnlyList<PartyInstallment> Reallocate(PartySnapshot snapshot,
        PartyInstallment installment, decimal amount)
    {
        var paid = snapshot.Payments.Where(payment => payment.InstallmentId == installment.Id)
            .Sum(payment => payment.Amount);
        var remaining = installment.Amount - paid;
        if (remaining <= 0)
            throw new ArgumentException("Esta parcela já está quitada.");

        var excess = amount - remaining;
        if (excess <= 0) return Array.Empty<PartyInstallment>();

        var future = snapshot.Installments
            .Where(candidate => candidate.Category == installment.Category &&
                candidate.ItemId == installment.ItemId &&
                (candidate.DueDate > installment.DueDate ||
                 candidate.DueDate == installment.DueDate &&
                 string.CompareOrdinal(candidate.Id, installment.Id) > 0))
            .OrderBy(candidate => candidate.DueDate).ThenBy(candidate => candidate.Id, StringComparer.Ordinal)
            .ToList();
        var futureBalances = future.Select(candidate => new
        {
            Installment = candidate,
            Remaining = Math.Max(0, candidate.Amount - snapshot.Payments
                .Where(payment => payment.InstallmentId == candidate.Id).Sum(payment => payment.Amount))
        }).ToList();
        if (excess > futureBalances.Sum(candidate => candidate.Remaining))
            throw new ArgumentException("O valor pago ultrapassa o saldo restante das parcelas deste item.");

        var updated = new List<PartyInstallment> { installment with { Amount = installment.Amount + excess } };
        foreach (var futureInstallment in futureBalances)
        {
            if (excess <= 0) break;
            var transfer = Math.Min(excess, futureInstallment.Remaining);
            if (transfer > 0)
                updated.Add(futureInstallment.Installment with
                { Amount = futureInstallment.Installment.Amount - transfer });
            excess -= transfer;
        }
        return updated;
    }
}
