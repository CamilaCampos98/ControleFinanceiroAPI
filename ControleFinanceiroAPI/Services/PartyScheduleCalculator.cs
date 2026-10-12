namespace ControleFinanceiroAPI.Services;

public static class PartyScheduleCalculator
{
    public static IReadOnlyList<decimal> Split(decimal remaining, int count, decimal minimum)
    {
        if (remaining <= 0 || count <= 0 || minimum < 0)
            throw new ArgumentException("Informe um saldo positivo, meses e parcela mínima válida.");
        if (Math.Round(remaining, 2) != remaining)
            throw new ArgumentException("O saldo deve ter no máximo duas casas decimais.");
        if (remaining < count * minimum)
            throw new ArgumentException($"O saldo não permite {count} parcelas de pelo menos {minimum:C2}.");

        var baseCents = (long)(remaining * 100) / count;
        var remainderCents = (long)(remaining * 100) % count;
        return Enumerable.Range(0, count)
            .Select(index => (baseCents + (index < remainderCents ? 1 : 0)) / 100m)
            .ToArray();
    }
}
