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

        var difference = amount - remaining;
        if (difference == 0) return Array.Empty<PartyInstallment>();

        var future = Future(snapshot.Installments, installment).ToList();
        if (difference < 0)
        {
            if (future.Count == 0) return Array.Empty<PartyInstallment>();
            return
            [
                installment with { Amount = paid + amount },
                future[0] with { Amount = future[0].Amount - difference }
            ];
        }

        var futureBalances = future.Select(candidate => new
        {
            Installment = candidate,
            Remaining = Math.Max(0, candidate.Amount - snapshot.Payments
                .Where(payment => payment.InstallmentId == candidate.Id).Sum(payment => payment.Amount))
        }).ToList();
        if (difference > futureBalances.Sum(candidate => candidate.Remaining))
            throw new ArgumentException("O valor pago ultrapassa o saldo restante das parcelas deste item.");

        var updated = new List<PartyInstallment> { installment with { Amount = installment.Amount + difference } };
        foreach (var futureInstallment in futureBalances)
        {
            if (difference <= 0) break;
            var transfer = Math.Min(difference, futureInstallment.Remaining);
            if (transfer > 0)
                updated.Add(futureInstallment.Installment with
                { Amount = futureInstallment.Installment.Amount - transfer });
            difference -= transfer;
        }
        return updated;
    }

    public static IReadOnlyList<PartyInstallment> ReallocateExisting(PartySnapshot snapshot)
    {
        var adjusted = snapshot.Installments.ToDictionary(item => item.Id);
        foreach (var installment in snapshot.Installments.OrderBy(item => item.DueDate)
                     .ThenBy(item => item.Id, StringComparer.Ordinal))
        {
            var current = adjusted[installment.Id];
            var paid = snapshot.Payments.Where(payment => payment.InstallmentId == current.Id)
                .Sum(payment => payment.Amount);
            if (paid <= 0 || paid >= current.Amount) continue;
            var next = Future(adjusted.Values, current).FirstOrDefault();
            if (next == null) continue;
            var difference = current.Amount - paid;
            adjusted[current.Id] = current with { Amount = paid };
            adjusted[next.Id] = next with { Amount = next.Amount + difference };
        }
        return adjusted.Values.Where(item => item.Amount != snapshot.Installments
            .First(original => original.Id == item.Id).Amount).ToList();
    }

    private static IEnumerable<PartyInstallment> Future(IEnumerable<PartyInstallment> installments,
        PartyInstallment current) => installments
        .Where(candidate => candidate.Category == current.Category && candidate.ItemId == current.ItemId &&
            (candidate.DueDate > current.DueDate || candidate.DueDate == current.DueDate &&
             string.CompareOrdinal(candidate.Id, current.Id) > 0))
        .OrderBy(candidate => candidate.DueDate).ThenBy(candidate => candidate.Id, StringComparer.Ordinal);
}
