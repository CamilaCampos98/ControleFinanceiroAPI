using ControleFinanceiroAPI.Models;
using ControleFinanceiroAPI.Services;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using System.Text.Json;

public partial class GoogleSheetsService
{
    private sealed record ClosingPlan(int RowIndex, DateTime Start, VacationPersonPreview Person,
        List<VacationPersonPreview> People);

    private sealed record ClosingFixedRow(int RowIndex, string Type, decimal Amount);

    private sealed record ClosingState(VacationPeriodClosingPreview Preview, List<ClosingPlan> Plans,
        List<ClosingFixedRow> GuardadoRows);

    public VacationPeriodClosingPreview PreviewVacationPeriodClosing(string person, string period) =>
        ReadVacationPeriodClosing(person, period).Preview;

    public VacationPeriodClosingPreview CloseVacationPeriod(VacationPeriodClosingRequest request)
    {
        lock (_lockId)
        {
            var state = ReadVacationPeriodClosing(request.Person, request.Period);
            var preview = state.Preview;
            if (!preview.CanClose)
                throw new InvalidOperationException(preview.Message ?? "Este período não pode ser fechado para férias.");
            if (preview.CurrentBalance != request.ExpectedBalance ||
                preview.ReductionNeeded != request.ExpectedReduction)
                throw new InvalidOperationException("Os valores mudaram. Atualize o resumo e confira o fechamento novamente.");

            var sheet = _service.Spreadsheets.Get(SpreadsheetId).Execute();
            var fixedSheetId = sheet.Sheets.First(item => item.Properties.Title == "Fixos").Properties.SheetId;
            var vacationSheetId = sheet.Sheets.First(item => item.Properties.Title == "Ferias").Properties.SheetId;
            var requests = new List<Request>();
            var remaining = preview.ReductionNeeded;
            var plans = state.Plans.OrderBy(plan => plan.Start).ToList();

            // Reduz primeiro o Guardado livre; depois a reserva da viagem; por último a renda protegida.
            foreach (var fixedRow in state.GuardadoRows.OrderBy(row => ClosingTypePriority(row.Type)).ThenBy(row => row.RowIndex))
            {
                if (remaining <= 0) break;
                var amount = Math.Min(remaining, fixedRow.Amount);
                if (amount <= 0) continue;
                if (ClosingTypePriority(fixedRow.Type) == 1)
                    AssignVacationReserveReduction(plans, amount, income: false);
                else if (ClosingTypePriority(fixedRow.Type) == 2)
                    AssignVacationReserveReduction(plans, amount, income: true);

                requests.Add(new Request { RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = fixedSheetId, StartRowIndex = fixedRow.RowIndex,
                        EndRowIndex = fixedRow.RowIndex + 1, StartColumnIndex = 5, EndColumnIndex = 6 },
                    Cell = new CellData { UserEnteredValue = new ExtendedValue
                    { NumberValue = (double)(fixedRow.Amount - amount) } },
                    Fields = "userEnteredValue"
                } });
                remaining -= amount;
            }
            if (remaining > 0)
                throw new InvalidOperationException("O Guardado disponível não cobre o ajuste necessário.");

            foreach (var plan in plans)
            {
                plan.Person.ClosedPeriods ??= new List<string>();
                if (!plan.Person.ClosedPeriods.Contains(request.Period, StringComparer.OrdinalIgnoreCase))
                    plan.Person.ClosedPeriods.Add(request.Period);
            }
            foreach (var group in plans.GroupBy(plan => plan.RowIndex))
            {
                var plan = group.First();
                requests.Add(new Request { RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = vacationSheetId, StartRowIndex = plan.RowIndex,
                        EndRowIndex = plan.RowIndex + 1, StartColumnIndex = 4, EndColumnIndex = 5 },
                    Cell = new CellData { UserEnteredValue = new ExtendedValue
                    { StringValue = JsonSerializer.Serialize(plan.People) } },
                    Fields = "userEnteredValue"
                } });
            }
            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, SpreadsheetId).Execute();
            return new VacationPeriodClosingPreview
            {
                HasPlannedVacation = true, IsClosed = true, CurrentBalance = preview.CurrentBalance,
                ReductionNeeded = preview.ReductionNeeded, AvailableGuardado = preview.AvailableGuardado,
                ProjectedBalance = preview.ProjectedBalance,
                TripReserveShortfall = plans.Sum(plan => plan.Person.TripReserveShortfall),
                IncomeReserveShortfall = plans.Sum(plan => plan.Person.IncomeReserveShortfall),
                Message = "Período fechado. As reservas de férias foram atualizadas."
            };
        }
    }

    private static int ClosingTypePriority(string type) => type.Trim().ToLowerInvariant() switch
    {
        "guardado" => 0,
        "guardado para férias" => 1,
        "guardado para dias sem faturamento" => 2,
        _ => 3
    };

    internal static decimal RequiredGuardadoReduction(decimal balance) => Math.Max(0, 200m - balance);

    private static void AssignVacationReserveReduction(List<ClosingPlan> plans, decimal amount, bool income)
    {
        var unmatched = ReducePlanReserves(plans.Select(plan => plan.Person).ToList(), amount, income);
        if (unmatched > 0)
            throw new InvalidOperationException("Não foi possível vincular toda a redução ao plano de férias. Confira as reservas antes de fechar.");
    }

    internal static decimal ReducePlanReserves(IReadOnlyList<VacationPersonPreview> people, decimal amount, bool income)
    {
        foreach (var person in people)
        {
            if (amount <= 0) break;
            var capacity = income
                ? Math.Max(0, person.IncomeCoverageByPeriod.Values.Sum())
                : Math.Max(0, person.BudgetShare - person.TripReserveShortfall);
            var allocated = Math.Min(amount, capacity);
            if (allocated <= 0) continue;
            if (income)
            {
                person.IncomeReserveShortfall += allocated;
                person.IncomeReserveFunded = Math.Max(0, person.IncomeReserveFunded - allocated);
                var remaining = allocated;
                foreach (var period in person.IncomeCoverageByPeriod.Keys
                    .OrderBy(value => DateTime.ParseExact(value, "MM/yyyy", CultureInfo.InvariantCulture)).ToList())
                {
                    var cut = Math.Min(remaining, person.IncomeCoverageByPeriod[period]);
                    person.IncomeCoverageByPeriod[period] -= cut;
                    remaining -= cut;
                    if (remaining <= 0) break;
                }
            }
            else person.TripReserveShortfall += allocated;
            amount -= allocated;
        }
        return amount;
    }

    private ClosingState ReadVacationPeriodClosing(string person, string period)
    {
        person = person?.Trim() ?? "";
        if (person.Length == 0 || !DateTime.TryParseExact(period, "MM/yyyy", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var periodDate))
            throw new ArgumentException("Informe uma pessoa e um período válido (MM/aaaa).");

        var planRows = VacationRows();
        var plans = new List<ClosingPlan>();
        var alreadyClosed = false;
        for (var index = 1; index < planRows.Count; index++)
        {
            var row = planRows[index];
            if (string.IsNullOrWhiteSpace(row.ElementAtOrDefault(0)?.ToString())) continue;
            var start = DateTime.ParseExact(row.ElementAtOrDefault(1)?.ToString() ?? "", "yyyy-MM-dd",
                CultureInfo.InvariantCulture);
            var people = JsonSerializer.Deserialize<List<VacationPersonPreview>>(row.ElementAtOrDefault(4)?.ToString() ?? "[]") ?? new();
            var participant = people.FirstOrDefault(item => string.Equals(item.Person, person, StringComparison.OrdinalIgnoreCase));
            if (participant == null) continue;
            alreadyClosed |= participant.ClosedPeriods?.Contains(period, StringComparer.OrdinalIgnoreCase) == true;
            if (start.Date >= VacationPlanningService.TodayBrazil &&
                VacationPlanningService.SalaryPeriod(start) > periodDate)
                plans.Add(new ClosingPlan(index, start, participant, people));
        }
        plans = plans.OrderBy(plan => plan.Start).ToList();

        var fixedRows = ReadData("Fixos!A1:H") ?? Array.Empty<IList<object>>();
        var guardado = fixedRows.Skip(1).Select((row, index) => new ClosingFixedRow(index + 1,
                row.ElementAtOrDefault(1)?.ToString()?.Trim() ?? "", ParseDecimal(row.ElementAtOrDefault(5)?.ToString())))
            .Where(item => item.Amount > 0 && ClosingTypePriority(item.Type) < 3 &&
                string.Equals(fixedRows[item.RowIndex].ElementAtOrDefault(2)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(fixedRows[item.RowIndex].ElementAtOrDefault(3)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var fixedTotal = fixedRows.Skip(1).Where(row =>
            string.Equals(row.ElementAtOrDefault(2)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(3)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(5)?.ToString()));
        var config = (ReadData("Config!A1:F") ?? Array.Empty<IList<object>>()).Skip(1).Where(row =>
            string.Equals(row.ElementAtOrDefault(0)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(3)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase)).ToList();
        var salary = config.Where(row => string.Equals(row.ElementAtOrDefault(1)?.ToString()?.Trim(), "Salario", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.ElementAtOrDefault(1)?.ToString()?.Trim(), "Salário", StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(2)?.ToString()));
        var extras = config.Sum(row => ParseDecimal(row.ElementAtOrDefault(5)?.ToString()));
        var purchases = (ReadData("Controle!A1:J") ?? Array.Empty<IList<object>>()).Skip(1).Where(row =>
            string.Equals(row.ElementAtOrDefault(5)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(7)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(4)?.ToString()));
        var (losses, coverage) = ReadVacationIncomeAdjustments();
        var key = (person.ToUpperInvariant(), period);
        var balance = Math.Max(0, salary - losses.GetValueOrDefault(key)) + extras + coverage.GetValueOrDefault(key)
            - fixedTotal - purchases;
        var reduction = RequiredGuardadoReduction(balance);
        var available = guardado.Sum(row => row.Amount);
        var hasPlan = plans.Count > 0;
        var periodFinished = periodDate.AddDays(24) <= VacationPlanningService.TodayBrazil;
        var canClose = hasPlan && !alreadyClosed && periodFinished && reduction <= available;
        var message = !hasPlan ? "Disponível apenas para quem tem férias programadas após este período."
            : alreadyClosed ? "Este período já foi fechado para as férias programadas."
            : !periodFinished ? "Aguarde o fim do período para conferir as faturas e fechá-lo."
            : reduction > available ? "O Guardado não é suficiente para manter R$ 200 de saldo neste período."
            : null;
        return new ClosingState(new VacationPeriodClosingPreview
        {
            HasPlannedVacation = hasPlan, IsClosed = alreadyClosed, CanClose = canClose,
            CurrentBalance = balance, ReductionNeeded = reduction, AvailableGuardado = available,
            ProjectedBalance = balance + Math.Min(reduction, available),
            TripReserveShortfall = plans.Sum(plan => plan.Person.TripReserveShortfall),
            IncomeReserveShortfall = plans.Sum(plan => plan.Person.IncomeReserveShortfall),
            Message = message
        }, plans, guardado);
    }
}
