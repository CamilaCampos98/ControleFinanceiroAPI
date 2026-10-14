using ControleFinanceiroAPI.Models;
using ControleFinanceiroAPI.Services;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using System.Text.Json;

public partial class GoogleSheetsService
{
    private sealed record ClosingPlan(int RowIndex, DateTime Start, VacationPersonPreview Person,
        List<VacationPersonPreview> People);

    private sealed record ClosingFixedRow(int RowIndex, string Id, string Type, decimal Amount);

    internal sealed record ClosingAdjustedRow(int RowIndex, string Id, decimal Amount);

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
            if (!SameReviewedCards(preview.CardsToReview, request.ReviewedCards))
                throw new InvalidOperationException("Confira todas as faturas dos cartões antes de fechar o período.");

            var sheet = _service.Spreadsheets.Get(SpreadsheetId).Execute();
            var fixedSheetId = sheet.Sheets.First(item => item.Properties.Title == "Fixos").Properties.SheetId;
            var vacationSheetId = state.Plans.Count > 0
                ? sheet.Sheets.First(item => item.Properties.Title == "Ferias").Properties.SheetId : null;
            var closingSheetId = sheet.Sheets.FirstOrDefault(item => item.Properties.Title == "FechamentosPeriodos")?.Properties.SheetId;
            if (closingSheetId == null)
            {
                var created = _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request> { new() { AddSheet = new AddSheetRequest
                    { Properties = new SheetProperties { Title = "FechamentosPeriodos" } } } }
                }, SpreadsheetId).Execute();
                closingSheetId = created.Replies[0].AddSheet.Properties.SheetId;
                var header = new ValueRange { Values = new List<IList<object>>
                { new List<object> { "Pessoa", "Periodo", "FechadoEm", "CartoesConferidos" } } };
                var update = _service.Spreadsheets.Values.Update(header, SpreadsheetId, "FechamentosPeriodos!A1:D1");
                update.ValueInputOption = Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                update.Execute();
            }
            var adjustmentRequests = new List<Request>();
            var adjustedRows = new List<ClosingAdjustedRow>();
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

                var adjustedAmount = fixedRow.Amount - amount;
                adjustedRows.Add(new ClosingAdjustedRow(fixedRow.RowIndex, fixedRow.Id, adjustedAmount));
                adjustmentRequests.Add(new Request { RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = fixedSheetId, StartRowIndex = fixedRow.RowIndex,
                        EndRowIndex = fixedRow.RowIndex + 1, StartColumnIndex = 5, EndColumnIndex = 6 },
                    Cell = new CellData { UserEnteredValue = new ExtendedValue
                    { NumberValue = (double)adjustedAmount } },
                    Fields = "userEnteredValue"
                } });
                remaining -= amount;
            }
            if (remaining > 0)
                throw new InvalidOperationException("O Guardado disponível não cobre o ajuste necessário.");

            foreach (var group in plans.GroupBy(plan => plan.RowIndex))
            {
                var plan = group.First();
                adjustmentRequests.Add(new Request { RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = vacationSheetId, StartRowIndex = plan.RowIndex,
                        EndRowIndex = plan.RowIndex + 1, StartColumnIndex = 4, EndColumnIndex = 5 },
                    Cell = new CellData { UserEnteredValue = new ExtendedValue
                    { StringValue = JsonSerializer.Serialize(plan.People) } },
                    Fields = "userEnteredValue"
                } });
            }
            if (adjustmentRequests.Count > 0)
                _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
                { Requests = adjustmentRequests }, SpreadsheetId).Execute();

            var actualFixedRows = ReadData("Fixos!A1:H") ?? Array.Empty<IList<object>>();
            var actualBalance = ReadVacationPeriodClosing(request.Person, request.Period).Preview.CurrentBalance;
            if (!ClosingAdjustmentMatches(actualFixedRows, adjustedRows, actualBalance, preview.ProjectedBalance))
                throw new InvalidOperationException("O ajuste do Guardado não foi confirmado na planilha. " +
                    "O período não foi marcado como fechado. Atualize o resumo antes de tentar novamente.");

            var closingRequests = new List<Request>();
            foreach (var plan in plans)
            {
                plan.Person.ClosedPeriods ??= new List<string>();
                if (!plan.Person.ClosedPeriods.Contains(request.Period, StringComparer.OrdinalIgnoreCase))
                    plan.Person.ClosedPeriods.Add(request.Period);
            }
            foreach (var group in plans.GroupBy(plan => plan.RowIndex))
            {
                var plan = group.First();
                closingRequests.Add(new Request { RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = vacationSheetId, StartRowIndex = plan.RowIndex,
                        EndRowIndex = plan.RowIndex + 1, StartColumnIndex = 4, EndColumnIndex = 5 },
                    Cell = new CellData { UserEnteredValue = new ExtendedValue
                    { StringValue = JsonSerializer.Serialize(plan.People) } },
                    Fields = "userEnteredValue"
                } });
            }
            closingRequests.Add(new Request { AppendCells = new AppendCellsRequest
            {
                SheetId = closingSheetId,
                Rows = new List<RowData> { new() { Values = new List<CellData>
                {
                    new() { UserEnteredValue = new ExtendedValue { StringValue = request.Person.Trim() } },
                    new() { UserEnteredValue = new ExtendedValue { StringValue = request.Period.Trim() } },
                    new() { UserEnteredValue = new ExtendedValue { StringValue = DateTimeOffset.UtcNow.ToString("O") } },
                    new() { UserEnteredValue = new ExtendedValue { StringValue = JsonSerializer.Serialize(preview.CardsToReview.Select(card => card.Name)) } }
                } } },
                Fields = "userEnteredValue"
            } });
            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            { Requests = closingRequests }, SpreadsheetId).Execute();
            var completed = ReadVacationPeriodClosing(request.Person, request.Period).Preview;
            if (!completed.IsClosed || completed.CurrentBalance != preview.ProjectedBalance)
                throw new InvalidOperationException("Não foi possível confirmar o fechamento na planilha. " +
                    "Atualize o resumo antes de tentar novamente.");
            return new VacationPeriodClosingPreview
            {
                HasPlannedVacation = preview.HasPlannedVacation, HasGuardado = preview.HasGuardado,
                IsClosed = true, CurrentBalance = preview.CurrentBalance,
                ReductionNeeded = preview.ReductionNeeded, AvailableGuardado = preview.AvailableGuardado,
                ProjectedBalance = preview.ProjectedBalance,
                TripReserveShortfall = plans.Sum(plan => plan.Person.TripReserveShortfall),
                IncomeReserveShortfall = plans.Sum(plan => plan.Person.IncomeReserveShortfall),
                CardsToReview = preview.CardsToReview,
                Message = preview.HasPlannedVacation
                    ? "Período fechado. As reservas de férias foram atualizadas."
                    : "Período fechado. O Guardado foi ajustado, se necessário."
            };
        }
    }

    internal static bool SameReviewedCards(IReadOnlyList<VacationPeriodClosingCard> cards, IReadOnlyList<string>? reviewed)
    {
        var expected = cards.Select(card => card.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var confirmed = (reviewed ?? Array.Empty<string>()).Select(name => (name ?? "").Trim())
            .Where(name => name.Length > 0).ToList();
        return confirmed.Count == expected.Count && confirmed.ToHashSet(StringComparer.OrdinalIgnoreCase).SetEquals(expected);
    }

    internal static bool ClosingAdjustmentMatches(IList<IList<object>> rows,
        IReadOnlyList<ClosingAdjustedRow> adjustedRows, decimal actualBalance, decimal projectedBalance) =>
        actualBalance == projectedBalance && adjustedRows.All(adjusted =>
            adjusted.RowIndex < rows.Count &&
            string.Equals(rows[adjusted.RowIndex].ElementAtOrDefault(0)?.ToString()?.Trim(),
                adjusted.Id, StringComparison.Ordinal) &&
            ParseClosingAmount(rows[adjusted.RowIndex].ElementAtOrDefault(5)?.ToString()) == adjusted.Amount);

    private static decimal ParseClosingAmount(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return 0;
        value = value.Replace("R$", "").Trim().Replace(" ", "");
        if (value.Contains(',')) value = value.Replace(".", "").Replace(',', '.');
        return decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount)
            ? amount : decimal.MinValue;
    }

    private static int ClosingTypePriority(string type) => type.Trim().ToLowerInvariant() switch
    {
        "guardado" => 0,
        "guardado para férias" => 1,
        "guardado para dias sem faturamento" => 2,
        _ => 3
    };

    internal static decimal RequiredGuardadoReduction(decimal balance) => Math.Max(0, 200m - balance);

    internal static bool PeriodFinished(DateTime periodDate, DateTime today) => periodDate.AddDays(25) <= today.Date;

    internal static bool CanClosePeriod(bool hasPlan, decimal available, bool alreadyClosed,
        bool periodFinished, decimal reduction) =>
        (hasPlan || available > 0) && !alreadyClosed && periodFinished && reduction <= available;

    internal static List<VacationPeriodClosingCard> ClosingCardsForPeriod(string person, DateTime periodDate,
        IList<IList<object>> configRows, IList<IList<object>> purchaseRows)
    {
        var names = new[] { "ITAU", "BRADESCO", "SANTANDER", "RIACHUELO", "C&A", "OUTROS" };
        var configs = configRows.Select((row, index) => new
            {
                Row = row,
                Index = index,
                Person = row.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "",
                Period = DateTime.TryParseExact(row.ElementAtOrDefault(3)?.ToString()?.Trim(), "MM/yyyy",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) ? date : DateTime.MinValue
            })
            .Where(item => string.Equals(item.Person, person, StringComparison.OrdinalIgnoreCase) &&
                item.Period != DateTime.MinValue)
            .OrderByDescending(item => item.Period <= periodDate)
            .ThenBy(item => Math.Abs((item.Period.Year - periodDate.Year) * 12 + item.Period.Month - periodDate.Month))
            .ThenByDescending(item => item.Index)
            .ToList();
        var cards = new Dictionary<string, VacationPeriodClosingCard>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < names.Length; index++)
        {
            var day = configs.Select(item => item.Row.ElementAtOrDefault(index + 6)?.ToString())
                .Select(value => int.TryParse(value, out var parsed) && parsed is >= 1 and <= 31 ? parsed : 0)
                .FirstOrDefault(value => value > 0);
            if (day == 0) continue;
            var nextMonth = periodDate.AddMonths(1);
            cards[names[index]] = new VacationPeriodClosingCard
            {
                Name = names[index],
                ClosingDate = new DateTime(nextMonth.Year, nextMonth.Month,
                    Math.Min(day, DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month)))
            };
        }
        foreach (var row in purchaseRows)
        {
            if (!string.Equals(row.ElementAtOrDefault(5)?.ToString()?.Trim(), periodDate.ToString("MM/yyyy"), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(row.ElementAtOrDefault(7)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
                continue;
            var name = RemoverPalavras(row.ElementAtOrDefault(9)?.ToString()?.Trim() ?? "").ToUpperInvariant();
            if (name.Length > 0 && !cards.ContainsKey(name))
                cards[name] = new VacationPeriodClosingCard { Name = name };
        }
        return cards.Values.OrderBy(card => card.ClosingDate ?? DateTime.MaxValue)
            .ThenBy(card => card.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

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
        var spreadsheet = _service.Spreadsheets.Get(SpreadsheetId).Execute();
        var closingRows = spreadsheet.Sheets.Any(item => item.Properties.Title == "FechamentosPeriodos")
            ? ReadData("FechamentosPeriodos!A1:D") ?? Array.Empty<IList<object>>() : Array.Empty<IList<object>>();
        if (spreadsheet.Sheets.Any(item => item.Properties.Title == "FechamentosPeriodos") &&
            (closingRows.Count == 0 ||
             closingRows[0].ElementAtOrDefault(0)?.ToString() != "Pessoa" ||
             closingRows[0].ElementAtOrDefault(1)?.ToString() != "Periodo" ||
             closingRows[0].ElementAtOrDefault(2)?.ToString() != "FechadoEm" ||
             closingRows[0].ElementAtOrDefault(3)?.ToString() != "CartoesConferidos"))
            throw new InvalidOperationException("A aba FechamentosPeriodos já existe com colunas diferentes das esperadas.");
        var alreadyClosed = closingRows.Skip(1).Any(row =>
            string.Equals(row.ElementAtOrDefault(0)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(1)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase));
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
                row.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "",
                row.ElementAtOrDefault(1)?.ToString()?.Trim() ?? "", ParseDecimal(row.ElementAtOrDefault(5)?.ToString())))
            .Where(item => item.Amount > 0 && ClosingTypePriority(item.Type) < 3 &&
                string.Equals(fixedRows[item.RowIndex].ElementAtOrDefault(2)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(fixedRows[item.RowIndex].ElementAtOrDefault(3)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var fixedTotal = fixedRows.Skip(1).Where(row =>
            string.Equals(row.ElementAtOrDefault(2)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(3)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(5)?.ToString()));
        var configRows = (ReadData("Config!A1:L") ?? Array.Empty<IList<object>>()).Skip(1).ToList();
        var config = configRows.Where(row =>
            string.Equals(row.ElementAtOrDefault(0)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(3)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase)).ToList();
        var salary = config.Where(row => string.Equals(row.ElementAtOrDefault(1)?.ToString()?.Trim(), "Salario", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.ElementAtOrDefault(1)?.ToString()?.Trim(), "Salário", StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(2)?.ToString()));
        var extras = config.Sum(row => ParseDecimal(row.ElementAtOrDefault(5)?.ToString()));
        var purchaseRows = (ReadData("Controle!A1:J") ?? Array.Empty<IList<object>>()).Skip(1).ToList();
        var purchases = purchaseRows.Where(row =>
            string.Equals(row.ElementAtOrDefault(5)?.ToString()?.Trim(), period, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(row.ElementAtOrDefault(7)?.ToString()?.Trim(), person, StringComparison.OrdinalIgnoreCase))
            .Sum(row => ParseDecimal(row.ElementAtOrDefault(4)?.ToString()));
        var cardsToReview = ClosingCardsForPeriod(person, periodDate, configRows, purchaseRows);
        var (losses, coverage) = ReadVacationIncomeAdjustments();
        var key = (person.ToUpperInvariant(), period);
        var balance = Math.Max(0, salary - losses.GetValueOrDefault(key)) + extras + coverage.GetValueOrDefault(key)
            - fixedTotal - purchases;
        var reduction = RequiredGuardadoReduction(balance);
        var available = guardado.Sum(row => row.Amount);
        var hasPlan = plans.Count > 0;
        var periodFinished = PeriodFinished(periodDate, VacationPlanningService.TodayBrazil);
        var canClose = CanClosePeriod(hasPlan, available, alreadyClosed, periodFinished, reduction);
        var message = alreadyClosed ? "Este período já foi fechado."
            : !hasPlan && available <= 0 ? "Disponível quando houver Guardado ou férias programadas."
            : !periodFinished ? "Aguarde o fim do período para conferir as faturas e fechá-lo."
            : reduction > available ? "O Guardado não é suficiente para manter R$ 200 de saldo neste período."
            : null;
        return new ClosingState(new VacationPeriodClosingPreview
        {
            HasPlannedVacation = hasPlan, HasGuardado = available > 0,
            IsClosed = alreadyClosed, CanClose = canClose,
            CurrentBalance = balance, ReductionNeeded = reduction, AvailableGuardado = available,
            ProjectedBalance = balance + Math.Min(reduction, available),
            TripReserveShortfall = plans.Sum(plan => plan.Person.TripReserveShortfall),
            IncomeReserveShortfall = plans.Sum(plan => plan.Person.IncomeReserveShortfall),
            CardsToReview = cardsToReview,
            Message = message
        }, plans, guardado);
    }
}
