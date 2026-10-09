using ControleFinanceiroAPI.Models;
using System.Globalization;
using System.Net.Http.Json;

namespace ControleFinanceiroAPI.Services;

public sealed class VacationPlanningService
{
    private readonly GoogleSheetsService _sheets;
    private static readonly HttpClient HolidayClient = new() { Timeout = TimeSpan.FromSeconds(8) };

    public VacationPlanningService(GoogleSheetsService sheets) => _sheets = sheets;

    internal static DateTime TodayBrazil => TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow,
        TimeZoneInfo.FindSystemTimeZoneById("America/Sao_Paulo")).Date;

    public async Task<VacationPlanPreview> PreviewAsync(VacationPlanRequest request)
    {
        Validate(request);
        var salaries = (_sheets.ReadData("Config!A1:F") ?? Array.Empty<IList<object>>()).Skip(1).ToList();
        var fixedCosts = (_sheets.ReadData("Fixos!A1:H") ?? Array.Empty<IList<object>>()).Skip(1).ToList();
        var purchases = (_sheets.ReadData("Controle!A1:J") ?? Array.Empty<IList<object>>()).Skip(1).ToList();
        var availableGuardado = AvailableGuardado(request, fixedCosts);
        var (holidays, holidayWarning) = await LoadHolidaysAsync(request.StartDate.Year, request.StartDate.AddMonths(6).Year + 1);
        var duration = (request.EndDate.Date - request.StartDate.Date).Days;
        // Derive the transfers from current sheet data, never from amounts posted by the browser.
        request.GuardadoTransfers = new();
        if (request.UseGuardado)
        {
            var withoutTransfers = CalculateOption(request, request.StartDate.Date, request.EndDate.Date,
                salaries, fixedCosts, purchases, holidays);
            request.GuardadoTransfers = CalculateAutomaticTransfers(availableGuardado, withoutTransfers.People);
            ValidateGuardadoTransfers(request, availableGuardado);
        }
        var options = new List<VacationOptionPreview>
        {
            CalculateOption(request, request.StartDate.Date, request.EndDate.Date,
                salaries, fixedCosts, purchases, holidays)
        };
        var alternativeRequest = new VacationPlanRequest
        {
            StartDate = request.StartDate, EndDate = request.EndDate, Budget = request.Budget,
            Participants = request.Participants
        };
        options.AddRange(Enumerable.Range(1, 6)
            .Select(offset => CalculateOption(alternativeRequest, request.StartDate.AddMonths(offset).Date,
                request.StartDate.AddMonths(offset).Date.AddDays(duration), salaries, fixedCosts, purchases, holidays)));
        var preview = BuildPreview(options, holidayWarning);
        preview.AvailableGuardado = availableGuardado;
        preview.GuardadoTransfers = request.GuardadoTransfers;
        return preview;
    }

    internal static VacationPlanPreview BuildPreview(List<VacationOptionPreview> options, bool holidayWarning)
    {
        var registeredOptions = options.Where(option => !option.HasEstimatedSalary)
            .OrderByDescending(option => option.LowestPeriodBalanceAfter)
            .ThenByDescending(option => option.BalanceAfter)
            .ToList();

        return new VacationPlanPreview
        {
            Selected = options[0],
            Recommended = registeredOptions.FirstOrDefault() ?? options[0],
            Alternatives = registeredOptions.Where(option => option != options[0]).ToList(),
            Warning = string.Join(" ", new[]
            {
                options[0].HasEstimatedSalary
                    ? "A data selecionada inclui período sem salário cadastrado e usa estimativa de horas úteis × último valor-hora. Confira antes de confirmar." : null,
                holidayWarning ? "A consulta de feriados falhou; foram usados os feriados nacionais e locais conhecidos. Confira os dias úteis." : null
            }.Where(message => message != null))
        };
    }

    public async Task<VacationOptionPreview> ConfirmAsync(VacationPlanRequest request)
    {
        var preview = (await PreviewAsync(request)).Selected;
        if (request.ExpectedLostIncome != preview.LostIncome ||
            request.ExpectedLowestPeriodBalanceAfter != preview.LowestPeriodBalanceAfter ||
            request.ExpectedGuardadoTransfers.Count != request.GuardadoTransfers.Count ||
            request.GuardadoTransfers.Any(transfer => !request.ExpectedGuardadoTransfers.Any(expected =>
                expected.Id == transfer.Id && expected.RowNumber == transfer.RowNumber &&
                expected.VacationAmount == transfer.VacationAmount &&
                expected.IncomeAmount == transfer.IncomeAmount)))
            throw new InvalidOperationException("Os valores mudaram desde a prévia. Compare as datas novamente antes de confirmar.");
        var existing = _sheets.ReadVacationPlans();
        foreach (var plan in existing)
        {
            if (plan.EndDate.Date < request.StartDate.Date || plan.StartDate.Date > request.EndDate.Date) continue;
            if (plan.Participants.Any(saved => request.Participants.Any(current =>
                string.Equals(saved.Person, current.Person, StringComparison.OrdinalIgnoreCase))))
                throw new InvalidOperationException("Uma das pessoas já possui férias planejadas nesse intervalo.");
        }

        _sheets.WriteVacationPlan(request, preview);
        return preview;
    }

    internal static VacationOptionPreview CalculateOption(VacationPlanRequest request, DateTime start, DateTime end,
        List<IList<object>> salaries, List<IList<object>> fixedCosts, List<IList<object>> purchases, HashSet<DateTime> holidays)
    {
        var days = Enumerable.Range(0, (end - start).Days + 1).Select(offset => start.AddDays(offset))
            .Where(day => day.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday && !holidays.Contains(day.Date))
            .ToList();
        var vacationPeriod = SalaryPeriod(start);
        var currentPeriod = SalaryPeriod(TodayBrazil);
        var savingPeriods = Math.Max(0, (vacationPeriod.Year - currentPeriod.Year) * 12 + vacationPeriod.Month - currentPeriod.Month);
        var affected = Enumerable.Range(0, (end - start).Days + 1).Select(offset => start.AddDays(offset))
            .GroupBy(SalaryPeriod)
            .ToDictionary(group => group.Key, group => group.Count(days.Contains));
        var people = new List<VacationPersonPreview>();
        var estimatedSalary = false;
        decimal balanceBefore = 0;
        decimal lostIncome = 0;
        decimal allocatedBudget = 0;
        var balanceByPeriod = affected.Keys.ToDictionary(period => period, _ => 0m);

        for (var index = 0; index < request.Participants.Count; index++)
        {
            var participant = request.Participants[index];
            var share = index == request.Participants.Count - 1
                ? request.Budget - allocatedBudget
                : Math.Round(request.Budget * participant.Percent / 100, 2);
            allocatedBudget += share;
            var rate = LatestRate(salaries, participant.Person, vacationPeriod);
            if (rate <= 0) throw new InvalidOperationException($"Não há valor-hora cadastrado para {participant.Person}.");
            var personTransfers = request.GuardadoTransfers
                .Where(transfer => transfer.RowNumber >= 2 &&
                    fixedCosts.ElementAtOrDefault(transfer.RowNumber - 2) is { } row &&
                    Same(row, 0, transfer.Id) && Same(row, 3, participant.Person) &&
                    Same(row, 2, transfer.Period)).ToList();
            var allocated = personTransfers.Sum(transfer => transfer.VacationAmount);
            var incomeAllocated = personTransfers.Sum(transfer => transfer.IncomeAmount);
            if (allocated > share)
                throw new ArgumentException($"O valor destinado do Guardado para {participant.Person} excede sua parte da viagem.");
            var person = new VacationPersonPreview
            {
                Person = participant.Person.Trim(), Percent = participant.Percent, BudgetShare = share,
                HourlyRate = rate, GuardadoAllocated = allocated, IncomeGuardadoAllocated = incomeAllocated,
                MonthlySaving = savingPeriods == 0 ? 0 : Math.Round((share - allocated) / savingPeriods, 2)
            };
            var personBalanceByPeriod = affected.Keys.ToDictionary(period => period, _ => 0m);

            foreach (var (period, dayCount) in affected)
            {
                var periodLabel = period.ToString("MM/yyyy", CultureInfo.InvariantCulture);
                var salaryRow = salaries.LastOrDefault(row => Same(row, 0, person.Person) &&
                    (Same(row, 1, "Salario") || Same(row, 1, "Salário")) && Same(row, 3, periodLabel));
                var sourceRow = salaryRow ?? LatestSalaryRow(salaries, person.Person, period);
                var periodRate = sourceRow == null ? 0 : Money(sourceRow, 4);
                if (periodRate <= 0) throw new InvalidOperationException($"Não há valor-hora para {person.Person} em {periodLabel}.");
                var sourcePeriod = salaryRow == null && sourceRow != null
                    ? DateTime.ParseExact(Value(sourceRow, 3), new[] { "M/yyyy", "MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None)
                    : period;
                var hoursPerDay = HoursPerDay(sourceRow!, sourcePeriod, holidays);
                var salary = salaryRow == null
                    ? BusinessDays(period.AddMonths(-1).AddDays(25), period.AddDays(24), holidays) * hoursPerDay * periodRate
                    : Money(salaryRow, 2);
                estimatedSalary |= salaryRow == null;
                var extras = salaryRow == null ? 0 : Money(salaryRow, 5);
                var fixedTotal = fixedCosts.Where(row => Same(row, 2, periodLabel) && Same(row, 3, person.Person))
                    .Sum(row => Money(row, 5));
                var purchasesTotal = purchases.Where(row => Same(row, 5, periodLabel) && Same(row, 7, person.Person))
                    .Sum(row => Money(row, 4));
                var balance = salary + extras - fixedTotal - purchasesTotal;
                balanceBefore += balance;
                person.BalanceBefore += balance;
                var loss = dayCount * hoursPerDay * periodRate;
                person.UnpaidHours += dayCount * hoursPerDay;
                balanceByPeriod[period] += balance - loss;
                personBalanceByPeriod[period] += balance - loss;
                person.LossByPeriod[periodLabel] = loss;
                person.LostIncome += loss;
                lostIncome += loss;
            }
            if (incomeAllocated > person.LostIncome)
                throw new ArgumentException($"O valor destinado do Guardado para cobrir os dias sem trabalho de {person.Person} excede a receita não faturada.");
            person.IncomeMonthlySaving = savingPeriods == 0 ? 0 :
                Math.Round((person.LostIncome - incomeAllocated) / savingPeriods, 2);
            person.IncomeReserveFunded = savingPeriods == 0 ? incomeAllocated : person.LostIncome;
            decimal coverageAllocated = 0;
            var lossPeriods = person.LossByPeriod.OrderBy(item =>
                DateTime.ParseExact(item.Key, "MM/yyyy", CultureInfo.InvariantCulture)).ToList();
            for (var lossIndex = 0; lossIndex < lossPeriods.Count; lossIndex++)
            {
                var loss = lossPeriods[lossIndex];
                var coverage = lossIndex == lossPeriods.Count - 1
                    ? person.IncomeReserveFunded - coverageAllocated
                    : person.LostIncome == 0 ? 0 :
                        Math.Round(loss.Value * person.IncomeReserveFunded / person.LostIncome, 2);
                person.IncomeCoverageByPeriod[loss.Key] = coverage;
                coverageAllocated += coverage;
            }
            // A viagem é provisionada nos períodos anteriores; não a cobrar de novo
            // no saldo do período das férias.
            person.BalanceAfter = person.BalanceBefore - person.LostIncome;
            var tightestPersonPeriod = personBalanceByPeriod.MinBy(entry => entry.Value);
            person.LowestPeriodBalanceAfter = tightestPersonPeriod.Value;
            person.TightestPeriod = tightestPersonPeriod.Key.ToString("MM/yyyy", CultureInfo.InvariantCulture);
            people.Add(person);
        }

        var tightestPeriod = balanceByPeriod.MinBy(entry => entry.Value);
        return new VacationOptionPreview
        {
            StartDate = start, EndDate = end, BusinessDays = days.Count,
            Budget = request.Budget, LostIncome = lostIncome, BalanceBefore = balanceBefore,
            BalanceAfter = balanceBefore - lostIncome,
            LowestPeriodBalanceAfter = tightestPeriod.Value,
            TightestPeriod = tightestPeriod.Key.ToString("MM/yyyy", CultureInfo.InvariantCulture),
            SavingPeriods = savingPeriods, HasEstimatedSalary = estimatedSalary, People = people
        };
    }

    private static int BusinessDays(DateTime start, DateTime end, HashSet<DateTime> holidays) =>
        Enumerable.Range(0, (end.Date - start.Date).Days + 1).Select(offset => start.Date.AddDays(offset))
            .Count(day => day.DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday && !holidays.Contains(day));

    private static IList<object>? LatestSalaryRow(List<IList<object>> salaries, string person, DateTime period) => salaries
        .Where(row => Same(row, 0, person) && (Same(row, 1, "Salario") || Same(row, 1, "Salário")) &&
            DateTime.TryParseExact(Value(row, 3), new[] { "M/yyyy", "MM/yyyy" }, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var rowPeriod) && rowPeriod <= period)
        .OrderByDescending(row => DateTime.ParseExact(Value(row, 3), new[] { "M/yyyy", "MM/yyyy" },
            CultureInfo.InvariantCulture, DateTimeStyles.None))
        .FirstOrDefault();

    private static decimal LatestRate(List<IList<object>> salaries, string person, DateTime period) =>
        LatestSalaryRow(salaries, person, period) is { } row ? Money(row, 4) : 0;

    private static int HoursPerDay(IList<object> salaryRow, DateTime period, HashSet<DateTime> holidays)
    {
        var rate = Money(salaryRow, 4);
        var businessDays = BusinessDays(period.AddMonths(-1).AddDays(25), period.AddDays(24), holidays);
        if (rate <= 0 || businessDays == 0) return 8;
        var implied = Money(salaryRow, 2) / rate / businessDays;
        var rounded = (int)Math.Round(implied, 0, MidpointRounding.AwayFromZero);
        return rounded is >= 1 and <= 24 && Math.Abs(implied - rounded) <= 0.15m ? rounded : 8;
    }

    private static decimal Money(IList<object> row, int index)
    {
        var value = Value(row, index).Replace("R$", "", StringComparison.OrdinalIgnoreCase).Trim().Replace(" ", "");
        if (value.Contains(',')) value = value.Replace(".", "").Replace(',', '.');
        return decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount) ? amount : 0;
    }
    private static string Value(IList<object> row, int index) => row.ElementAtOrDefault(index)?.ToString()?.Trim() ?? "";
    private static bool Same(IList<object> row, int index, string value) =>
        string.Equals(Value(row, index), value, StringComparison.OrdinalIgnoreCase);

    internal static List<VacationGuardadoOption> AvailableGuardado(VacationPlanRequest request,
        List<IList<object>> fixedCosts)
    {
        var vacationPeriod = SalaryPeriod(request.StartDate);
        var currentPeriod = SalaryPeriod(TodayBrazil);
        return fixedCosts
            .Select((row, index) => new { Row = row, RowNumber = index + 2 })
            .Where(item => Same(item.Row, 1, "Guardado") &&
                request.Participants.Any(person => Same(item.Row, 3, person.Person)) &&
                DateTime.TryParseExact(Value(item.Row, 2), new[] { "M/yyyy", "MM/yyyy" },
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var period) &&
                period >= currentPeriod && period < vacationPeriod &&
                Money(item.Row, 5) > 0 && !string.IsNullOrWhiteSpace(Value(item.Row, 0)))
            .Select(item => new VacationGuardadoOption
            {
                Id = Value(item.Row, 0), RowNumber = item.RowNumber,
                Person = Value(item.Row, 3), Period = Value(item.Row, 2), Available = Money(item.Row, 5)
            })
            .OrderBy(option => DateTime.ParseExact(option.Period, new[] { "M/yyyy", "MM/yyyy" },
                CultureInfo.InvariantCulture, DateTimeStyles.None))
            .ThenBy(option => option.Person)
            .ToList();
    }

    internal static void ValidateGuardadoTransfers(VacationPlanRequest request,
        List<VacationGuardadoOption> available)
    {
        if (request.GuardadoTransfers.Count > 100 ||
            request.GuardadoTransfers.Any(transfer => string.IsNullOrWhiteSpace(transfer.Id) ||
                transfer.VacationAmount < 0 || transfer.IncomeAmount < 0 ||
                transfer.VacationAmount + transfer.IncomeAmount <= 0) ||
            request.GuardadoTransfers.Select(transfer => transfer.RowNumber).Distinct().Count() !=
                request.GuardadoTransfers.Count)
            throw new ArgumentException("Revise os valores escolhidos do Guardado.");
        foreach (var transfer in request.GuardadoTransfers)
        {
            var source = available.SingleOrDefault(option => option.RowNumber == transfer.RowNumber &&
                option.Id == transfer.Id &&
                string.Equals(option.Person, transfer.Person, StringComparison.OrdinalIgnoreCase) &&
                option.Period == transfer.Period);
            if (source == null || transfer.VacationAmount + transfer.IncomeAmount > source.Available)
                throw new InvalidOperationException("Um valor do Guardado mudou ou não está disponível. Compare as datas novamente.");
        }
    }

    internal static List<VacationGuardadoTransfer> CalculateAutomaticTransfers(
        List<VacationGuardadoOption> available, List<VacationPersonPreview> people)
    {
        var transfers = new List<VacationGuardadoTransfer>();
        foreach (var person in people)
        {
            var vacationRemaining = person.BudgetShare;
            var incomeRemaining = person.LostIncome;
            var sources = available.Where(option =>
                string.Equals(option.Person, person.Person, StringComparison.OrdinalIgnoreCase))
                .Select(option => new
                {
                    Source = option,
                    Limit = decimal.Floor(option.Available * .70m * 100m) / 100m
                })
                .Where(item => item.Limit > 0)
                .ToList();
            if (sources.Count == 0) continue;
            var totalLimit = sources.Sum(item => item.Limit);
            var target = Math.Min(totalLimit,
                decimal.Floor((vacationRemaining + incomeRemaining) * 100m) / 100m);
            if (target <= 0) continue;

            // Divide the needed amount across all eligible months in proportion to
            // their Guardado. Limits are rounded down so at least 30% remains.
            var shares = sources.Select(item => target * item.Limit / totalLimit).ToArray();
            var amounts = shares.Select(share => decimal.Floor(share * 100m) / 100m).ToArray();
            var remainingCents = (int)((target - amounts.Sum()) * 100m);
            foreach (var index in Enumerable.Range(0, sources.Count)
                .OrderByDescending(index => shares[index] - amounts[index]).ThenBy(index => index))
            {
                if (remainingCents == 0) break;
                if (amounts[index] + .01m > sources[index].Limit) continue;
                amounts[index] += .01m;
                remainingCents--;
            }
            if (remainingCents != 0)
                throw new InvalidOperationException("Não foi possível distribuir o Guardado entre os meses.");

            for (var index = 0; index < sources.Count; index++)
            {
                var source = sources[index].Source;
                var allowance = amounts[index];
                if (allowance <= 0) continue;
                var vacation = incomeRemaining == 0 ? allowance :
                    Math.Min(vacationRemaining, Math.Round(allowance * vacationRemaining /
                        (vacationRemaining + incomeRemaining), 2, MidpointRounding.AwayFromZero));
                var income = allowance - vacation;
                transfers.Add(new VacationGuardadoTransfer
                {
                    Id = source.Id, RowNumber = source.RowNumber, Person = source.Person,
                    Period = source.Period, VacationAmount = vacation, IncomeAmount = income
                });
                vacationRemaining -= vacation;
                incomeRemaining -= income;
            }
        }
        return transfers;
    }
    internal static DateTime SalaryPeriod(DateTime date) => new DateTime(date.Year, date.Month, 1).AddMonths(date.Day >= 26 ? 1 : 0);

    internal static void Validate(VacationPlanRequest request)
    {
        if (request.StartDate.Date < TodayBrazil || request.EndDate.Date < request.StartDate.Date ||
            request.EndDate.Date > request.StartDate.Date.AddDays(60) || request.StartDate.Date > TodayBrazil.AddMonths(18))
            throw new ArgumentException("Escolha datas futuras, com até 60 dias de férias e início nos próximos 18 meses.");
        if (request.Budget <= 0 || request.Budget > 10000000)
            throw new ArgumentException("Informe um orçamento total maior que zero.");
        if (request.Participants.Count == 0 || request.Participants.Count > 6 ||
            request.Participants.Any(person => string.IsNullOrWhiteSpace(person.Person) || person.Percent < 0) ||
            request.Participants.Select(person => person.Person.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count() != request.Participants.Count ||
            Math.Abs(request.Participants.Sum(person => person.Percent) - 100) > 0.01m)
            throw new ArgumentException("Selecione de 1 a 6 pessoas com percentuais que somem 100%.");
    }

    private static async Task<(HashSet<DateTime> Dates, bool Warning)> LoadHolidaysAsync(int firstYear, int lastYear)
    {
        var result = new HashSet<DateTime>();
        var warning = false;
        var years = Enumerable.Range(firstYear, lastYear - firstYear + 1).ToArray();
        var fetched = await Task.WhenAll(years.Select(async year =>
        {
            try { return await HolidayClient.GetFromJsonAsync<List<HolidayItem>>($"https://brasilapi.com.br/api/feriados/v1/{year}"); }
            catch (Exception exception) when (exception is HttpRequestException or TaskCanceledException) { return null; }
        }));
        for (var index = 0; index < years.Length; index++)
        {
            var year = years[index];
            var list = fetched[index];
            if (list == null) warning = true;
            else foreach (var item in list)
                if (DateTime.TryParse(item.Date, out var date)) result.Add(date.Date);
            foreach (var (month, day) in new[] { (1, 1), (4, 21), (5, 1), (9, 7), (10, 12),
                (11, 2), (11, 15), (11, 20), (12, 25) }) result.Add(new DateTime(year, month, day));
            result.Add(new DateTime(year, 7, 9)); // SP, adotado pela calculadora de horas úteis do portal.
            result.Add(new DateTime(year, 3, 26)); // Barueri, adotado pela calculadora de horas úteis do portal.
            var easter = Easter(year);
            result.Add(easter.AddDays(-48)); // Segunda-feira de Carnaval.
            result.Add(easter.AddDays(-47)); // Carnaval.
            result.Add(easter.AddDays(-2));  // Sexta-feira Santa.
            result.Add(easter.AddDays(60));  // Corpus Christi.
        }
        return (result, warning);
    }

    private static DateTime Easter(int year)
    {
        var a = year % 19; var b = year / 100; var c = year % 100;
        var d = b / 4; var e = b % 4; var f = (b + 8) / 25; var g = (b - f + 1) / 3;
        var h = (19 * a + b - d - g + 15) % 30; var i = c / 4; var k = c % 4;
        var l = (32 + 2 * e + 2 * i - h - k) % 7;
        var m = (a + 11 * h + 22 * l) / 451;
        var month = (h + l - 7 * m + 114) / 31;
        var day = ((h + l - 7 * m + 114) % 31) + 1;
        return new DateTime(year, month, day);
    }

    private sealed class HolidayItem
    {
        public string Date { get; set; } = "";
    }
}
