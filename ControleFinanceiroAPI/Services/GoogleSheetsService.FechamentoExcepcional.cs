public partial class GoogleSheetsService
{
    public static (DateTime inicio, DateTime fim) CalcularPeriodoFatura(
        DateTime mesReferencia, int diaFechamentoPadrao, string cartaoBase, string pessoa)
    {
        var fechamentoAtual = DataFechamento(mesReferencia, diaFechamentoPadrao, cartaoBase, pessoa);
        var fechamentoSeguinte = DataFechamento(mesReferencia.AddMonths(1), diaFechamentoPadrao,
            cartaoBase, pessoa);
        return (fechamentoAtual.AddDays(1), fechamentoSeguinte);
    }

    private static DateTime DataFechamento(DateTime mes, int diaPadrao, string cartaoBase, string pessoa)
    {
        // A fatura do Infinite fechou em 07/10/2026 para titular e adicional.
        // A configuração habitual continua sendo dia 8 nos demais fechamentos.
        var dia = cartaoBase.Equals("ITAU", StringComparison.OrdinalIgnoreCase) &&
            (pessoa.Equals("Camila", StringComparison.OrdinalIgnoreCase) ||
             pessoa.Equals("Diego", StringComparison.OrdinalIgnoreCase)) &&
            mes.Year == 2026 && mes.Month == 10 ? 7 : diaPadrao;
        return new DateTime(mes.Year, mes.Month, Math.Min(dia, DateTime.DaysInMonth(mes.Year, mes.Month)));
    }
}
