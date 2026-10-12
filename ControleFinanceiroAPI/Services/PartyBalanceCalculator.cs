using ControleFinanceiroAPI.Models;

namespace ControleFinanceiroAPI.Services;

public static class PartyBalanceCalculator
{
    public static decimal Unplanned(PartySnapshot snapshot, string category, string? itemId)
    {
        var items = snapshot.Items.Where(item => item.Category == category &&
            (itemId == null || item.Id == itemId));
        var knownCost = items.Sum(item => item.Amount ?? 0);
        var historicalPayments = snapshot.Payments.Where(payment => payment.InstallmentId == null &&
            (itemId == null ? payment.Category == category : payment.ItemId == itemId))
            .Sum(payment => payment.Amount);
        var planned = snapshot.Installments.Where(installment =>
            itemId == null ? installment.Category == category : installment.ItemId == itemId)
            .Sum(installment => installment.Amount);
        return knownCost - historicalPayments - planned;
    }
}
