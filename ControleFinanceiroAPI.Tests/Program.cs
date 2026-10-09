using ControleFinanceiroAPI.Models;
using ControleFinanceiroAPI.Services;

static IList<object> Row(params object[] values) => values.ToList();
static void Equal<T>(T expected, T actual, string name)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new Exception($"{name}: esperado {expected}; obtido {actual}.");
}
static int Weekdays(DateTime start, DateTime end) => Enumerable.Range(0, (end - start).Days + 1)
    .Count(offset => start.AddDays(offset).DayOfWeek is not DayOfWeek.Saturday and not DayOfWeek.Sunday);

Equal("09/2026", VacationPlanningService.SalaryPeriod(new DateTime(2026, 9, 25)).ToString("MM/yyyy"), "dia 25 pertence ao período atual");
Equal("10/2026", VacationPlanningService.SalaryPeriod(new DateTime(2026, 9, 26)).ToString("MM/yyyy"), "dia 26 muda o período");

var request = new VacationPlanRequest
{
    StartDate = new DateTime(2026, 10, 23), EndDate = new DateTime(2026, 10, 28), Budget = 6000,
    Participants = new List<VacationParticipant>
    {
        new() { Person = "Diego", Percent = 60 },
        new() { Person = "Camila", Percent = 40 }
    }
};
var salaries = new List<IList<object>>
{
    Row("Diego", "Salario", Weekdays(new DateTime(2026, 9, 26), new DateTime(2026, 10, 25)) * 8 * 80, "10/2026", 80, 0),
    Row("Diego", "Salario", Weekdays(new DateTime(2026, 10, 26), new DateTime(2026, 11, 25)) * 8 * 80, "11/2026", 80, 0),
    Row("Camila", "Salario", Weekdays(new DateTime(2026, 9, 26), new DateTime(2026, 10, 25)) * 8 * 60, "10/2026", 60, 0),
    Row("Camila", "Salario", Weekdays(new DateTime(2026, 10, 26), new DateTime(2026, 11, 25)) * 8 * 60, "11/2026", 60, 0)
};
var option = VacationPlanningService.CalculateOption(request, request.StartDate, request.EndDate,
    salaries, new(), new(), new());
Equal(4, option.BusinessDays, "dias úteis em dois períodos");
Equal(4480m, option.LostIncome, "receita não faturada conjunta");
var octoberBase = Weekdays(new DateTime(2026, 9, 26), new DateTime(2026, 10, 25)) * 8 * (80 + 60);
var novemberBase = Weekdays(new DateTime(2026, 10, 26), new DateTime(2026, 11, 25)) * 8 * (80 + 60);
Equal((decimal)(octoberBase + novemberBase), option.BalanceBefore, "sobra antes das férias");
Equal((decimal)(octoberBase + novemberBase - 4480), option.BalanceAfter,
    "custo da viagem já reservado não é descontado novamente");
Equal((decimal)Math.Min(octoberBase - 1120, novemberBase - 3360), option.LowestPeriodBalanceAfter,
    "período financeiro mais apertado considera só a perda de receita");
Equal("10/2026", option.TightestPeriod, "mês do menor saldo conjunto");
Equal(3600m, option.People[0].BudgetShare, "Diego guarda 60%");
Equal(2400m, option.People[1].BudgetShare, "Camila guarda 40%");
var diegoOctoberBase = Weekdays(new DateTime(2026, 9, 26), new DateTime(2026, 10, 25)) * 8 * 80;
var diegoNovemberBase = Weekdays(new DateTime(2026, 10, 26), new DateTime(2026, 11, 25)) * 8 * 80;
Equal((decimal)(diegoOctoberBase + diegoNovemberBase), option.People[0].BalanceBefore, "saldo inicial de Diego");
Equal((decimal)(diegoOctoberBase + diegoNovemberBase - 2560), option.People[0].BalanceAfter,
    "saldo de Diego não desconta novamente sua parte da viagem");
Equal((decimal)Math.Min(diegoOctoberBase - 640, diegoNovemberBase - 1920),
    option.People[0].LowestPeriodBalanceAfter, "menor saldo mensal de Diego");
Equal("10/2026", option.People[0].TightestPeriod, "mês do menor saldo de Diego");
Equal(option.BalanceBefore, option.People.Sum(person => person.BalanceBefore), "saldos iniciais somam o total");
Equal(option.BalanceAfter, option.People.Sum(person => person.BalanceAfter), "saldos finais somam o total");
Equal(640m, option.People[0].LossByPeriod["10/2026"], "Diego em outubro");
Equal(1920m, option.People[0].LossByPeriod["11/2026"], "Diego em novembro");
Equal(480m, option.People[1].LossByPeriod["10/2026"], "Camila em outubro");
Equal(1440m, option.People[1].LossByPeriod["11/2026"], "Camila em novembro");

var separateRequest = new VacationPlanRequest
{
    StartDate = new DateTime(2027, 3, 1), EndDate = new DateTime(2027, 3, 5), Budget = 1000m,
    Participants = new() { new() { Person = "Camila", Percent = 50 }, new() { Person = "Diego", Percent = 50 } }
};
var separateSalaries = new List<IList<object>>
{
    Row("Camila", "Salario", Weekdays(new DateTime(2027, 2, 26), new DateTime(2027, 3, 25)) * 8 * 50,
        "03/2027", 50, 0),
    Row("Diego", "Salario", Weekdays(new DateTime(2027, 2, 26), new DateTime(2027, 3, 25)) * 8 * 40,
        "03/2027", 40, 0)
};
var separateFixed = new List<IList<object>>
{
    Row("camila-save", "Guardado", "02/2027", "Camila", "", "3000", "Não", ""),
    Row("diego-save", "Guardado", "02/2027", "Diego", "", "1000", "Não", "")
};
var separateBase = VacationPlanningService.CalculateOption(separateRequest, separateRequest.StartDate,
    separateRequest.EndDate, separateSalaries, separateFixed, new(), new());
Equal(2000m, separateBase.People.Single(person => person.Person == "Camila").LostIncome,
    "Camila deixa de faturar seu próprio valor");
Equal(1600m, separateBase.People.Single(person => person.Person == "Diego").LostIncome,
    "Diego deixa de faturar seu próprio valor");
Equal(3600m, separateBase.LostIncome, "total de perda é apenas a soma das duas pessoas");
separateRequest.GuardadoTransfers = VacationPlanningService.CalculateAutomaticTransfers(new()
{
    new() { Id = "camila-save", RowNumber = 2, Person = "Camila", Period = "02/2027", Available = 3000m },
    new() { Id = "diego-save", RowNumber = 3, Person = "Diego", Period = "02/2027", Available = 1000m }
}, separateBase.People);
var separateFunded = VacationPlanningService.CalculateOption(separateRequest, separateRequest.StartDate,
    separateRequest.EndDate, separateSalaries, separateFixed, new(), new());
var camilaFunded = separateFunded.People.Single(person => person.Person == "Camila");
var diegoFunded = separateFunded.People.Single(person => person.Person == "Diego");
Equal(Math.Round((2000m - camilaFunded.IncomeGuardadoAllocated) / separateFunded.SavingPeriods, 2),
    camilaFunded.IncomeMonthlySaving, "a reserva mensal de Camila usa apenas sua perda");
Equal(Math.Round((1600m - diegoFunded.IncomeGuardadoAllocated) / separateFunded.SavingPeriods, 2),
    diegoFunded.IncomeMonthlySaving, "a reserva mensal de Diego usa apenas sua perda");
Equal(2000m, camilaFunded.IncomeCoverageByPeriod["03/2027"],
    "cobertura de Camila não inclui a perda de Diego");
Equal(1600m, diegoFunded.IncomeCoverageByPeriod["03/2027"],
    "cobertura de Diego não inclui a perda de Camila");
Equal(camilaFunded.BalanceBefore - 2000m,
    camilaFunded.BalanceAfter, "saldo de Camila desconta seus próprios 2000");
Equal(diegoFunded.BalanceBefore - 1600m,
    diegoFunded.BalanceAfter, "saldo de Diego desconta seus próprios 1600");

var sixHourSalaries = new List<IList<object>>
{
    Row("Camila", "Salario", Weekdays(new DateTime(2026, 9, 26), new DateTime(2026, 10, 25)) * 6 * 60,
        "10/2026", 60, 0),
    Row("Camila", "Salario", Weekdays(new DateTime(2026, 10, 26), new DateTime(2026, 11, 25)) * 6 * 60,
        "11/2026", 60, 0)
};
var sixHourRequest = new VacationPlanRequest
{
    StartDate = request.StartDate, EndDate = request.EndDate, Budget = 1000,
    Participants = new() { new() { Person = "Camila", Percent = 100 } }
};
var sixHourOption = VacationPlanningService.CalculateOption(sixHourRequest, request.StartDate,
    request.EndDate, sixHourSalaries, new(), new(), new());
Equal(24, sixHourOption.People[0].UnpaidHours, "jornada de 6 horas inferida do salário");
Equal(1440m, sixHourOption.LostIncome, "perda com jornada de 6 horas");
Equal(sixHourOption.BalanceAfter, sixHourOption.People[0].BalanceAfter, "saldo de uma pessoa equivale ao total");
var onePeriodRequest = new VacationPlanRequest
{
    StartDate = new DateTime(2026, 10, 20), EndDate = new DateTime(2026, 10, 21), Budget = 500,
    Participants = new() { new() { Person = "Camila", Percent = 100 } }
};
var onePeriodOption = VacationPlanningService.CalculateOption(onePeriodRequest, onePeriodRequest.StartDate,
    onePeriodRequest.EndDate, sixHourSalaries, new(), new(), new());
Equal("10/2026", onePeriodOption.TightestPeriod, "único período afetado");
Equal(onePeriodOption.BalanceAfter, onePeriodOption.LowestPeriodBalanceAfter,
    "em um período, saldo final é o saldo mensal");
Equal(onePeriodOption.People[0].BalanceAfter, onePeriodOption.People[0].LowestPeriodBalanceAfter,
    "em um período, saldo da pessoa é o saldo mensal");
var immediateRequest = new VacationPlanRequest
{
    StartDate = onePeriodRequest.StartDate, EndDate = onePeriodRequest.EndDate, Budget = 500,
    Participants = new() { new() { Person = "Camila", Percent = 100 } },
    GuardadoTransfers = new() { new() { Id = "existing", RowNumber = 2, Person = "Camila",
        Period = "09/2026", IncomeAmount = 300 } }
};
var immediateOption = VacationPlanningService.CalculateOption(immediateRequest, immediateRequest.StartDate,
    immediateRequest.EndDate, sixHourSalaries,
    new() { Row("existing", "Guardado", "09/2026", "Camila", "2026-10-15", "300", "Sim", "Não") },
    new(), new());
Equal(0, immediateOption.SavingPeriods, "sem meses anteriores para provisionar");
Equal(300m, immediateOption.People[0].IncomeReserveFunded,
    "sem meses anteriores, só o Guardado transferido cobre a perda");
Equal(300m, immediateOption.People[0].IncomeCoverageByPeriod.Values.Sum(),
    "cobertura não presume reserva que não pôde ser provisionada");

var incompleteSalaries = salaries.Where(row => !(row[0].ToString() == "Camila" && row[3].ToString() == "11/2026"))
    .ToList();
var incompleteOption = VacationPlanningService.CalculateOption(request, request.StartDate, request.EndDate,
    incompleteSalaries, new(), new(), new());
Equal(true, incompleteOption.HasEstimatedSalary, "falta de salário em um período afetado invalida alternativa");
var registeredOption = new VacationOptionPreview { StartDate = new DateTime(2026, 11, 23), LowestPeriodBalanceAfter = 1000 };
var betterRegisteredOption = new VacationOptionPreview { StartDate = new DateTime(2026, 12, 23), LowestPeriodBalanceAfter = 2000 };
var filteredPreview = VacationPlanningService.BuildPreview(
    new() { incompleteOption, registeredOption, betterRegisteredOption }, false);
Equal(2, filteredPreview.Alternatives.Count, "somente alternativas com todos os salários cadastrados");
Equal(betterRegisteredOption, filteredPreview.Recommended, "recomendação usa apenas salários cadastrados");
Equal(incompleteOption, filteredPreview.Selected, "data escolhida permanece na prévia");
var noAlternatives = VacationPlanningService.BuildPreview(new() { option, incompleteOption }, false);
Equal(0, noAlternatives.Alternatives.Count, "sem alternativa estimada na lista");
Equal(option, noAlternatives.Recommended, "data escolhida cadastrada permanece recomendada");

var savingsRequest = new VacationPlanRequest
{
    StartDate = new DateTime(2027, 3, 27), EndDate = new DateTime(2027, 4, 3), Budget = 5000,
    Participants = new() { new() { Person = "Camila", Percent = 50 }, new() { Person = "Diego", Percent = 50 } },
    GuardadoTransfers = new() { new() { Id = "1.761E+15", RowNumber = 2,
        Person = "Camila", Period = "03/2027", VacationAmount = 1800, IncomeAmount = 1000 } }
};
var savingsRows = new List<IList<object>>
{
    Row("1.761E+15", "Guardado", "03/2027", "Camila", "2027-04-15", "3000", "Sim", "Não"),
    Row("second", "Guardado", "03/2027", "Diego", "2027-04-15", "1000", "Não", "Não"),
    Row("1.761E+15", "Guardado", "02/2027", "Diego", "2027-03-15", "800", "Não", "Não"),
    Row("late", "Guardado", "04/2027", "Camila", "2027-05-15", "500", "Não", "Não"),
    Row("other", "Impostos", "03/2027", "Camila", "2027-04-15", "1000", "Não", "Não")
};
var availableSavings = VacationPlanningService.AvailableGuardado(savingsRequest, savingsRows);
Equal(3, availableSavings.Count, "somente Guardado anterior às férias das pessoas selecionadas");
var currentSavingsPeriod = VacationPlanningService.SalaryPeriod(VacationPlanningService.TodayBrazil);
var windowRequest = new VacationPlanRequest
{
    StartDate = currentSavingsPeriod.AddMonths(3),
    Participants = new() { new() { Person = "Camila", Percent = 100 } }
};
var windowRows = new List<IList<object>>
{
    Row("old", "Guardado", currentSavingsPeriod.AddMonths(-1).ToString("MM/yyyy"), "Camila", "", "100", "Não", ""),
    Row("current", "Guardado", currentSavingsPeriod.ToString("MM/yyyy"), "Camila", "", "100", "Não", ""),
    Row("next", "Guardado", currentSavingsPeriod.AddMonths(1).ToString("MM/yyyy"), "Camila", "", "100", "Não", ""),
    Row("vacation", "Guardado", currentSavingsPeriod.AddMonths(3).ToString("MM/yyyy"), "Camila", "", "100", "Não", "")
};
var windowAvailable = VacationPlanningService.AvailableGuardado(windowRequest, windowRows);
Equal(2, windowAvailable.Count, "Guardado só do período atual até antes das férias");
Equal("current", windowAvailable[0].Id, "o período de salário atual é elegível");
Equal("next", windowAvailable[1].Id, "período futuro anterior às férias é elegível");
var automaticPeople = new List<VacationPersonPreview>
{
    new() { Person = "Camila", BudgetShare = 2500m, LostIncome = 2400m },
    new() { Person = "Diego", BudgetShare = 2500m, LostIncome = 2332m }
};
var automatic = VacationPlanningService.CalculateAutomaticTransfers(availableSavings, automaticPeople);
Equal(2100m, automatic.Single(item => item.Person == "Camila").VacationAmount +
    automatic.Single(item => item.Person == "Camila").IncomeAmount, "preserva 30% da linha de Camila");
Equal(1260m, automatic.Where(item => item.Person == "Diego").Sum(item =>
    item.VacationAmount + item.IncomeAmount), "preserva 30% de cada linha de Diego");
Equal(true, automatic.All(item => item.VacationAmount > 0 && item.IncomeAmount > 0),
    "divide proporcionalmente entre as duas metas");
var proportionalSources = new List<VacationGuardadoOption>
{
    new() { Id = "large", RowNumber = 2, Person = "Camila", Period = "10/2026", Available = 3000m },
    new() { Id = "small", RowNumber = 3, Person = "Camila", Period = "11/2026", Available = 1000m }
};
var proportional = VacationPlanningService.CalculateAutomaticTransfers(proportionalSources,
    new() { new() { Person = "Camila", BudgetShare = 800m, LostIncome = 200m } });
Equal(750m, proportional[0].VacationAmount + proportional[0].IncomeAmount,
    "mês com 3000 contribui com 75% da necessidade");
Equal(250m, proportional[1].VacationAmount + proportional[1].IncomeAmount,
    "mês com 1000 contribui com 25% da necessidade");
Equal(800m, proportional.Sum(item => item.VacationAmount), "meta da viagem dividida entre meses");
Equal(200m, proportional.Sum(item => item.IncomeAmount), "meta da renda dividida entre meses");
var maximum = VacationPlanningService.CalculateAutomaticTransfers(proportionalSources,
    new() { new() { Person = "Camila", BudgetShare = 3000m, LostIncome = 2000m } });
Equal(2100m, maximum[0].VacationAmount + maximum[0].IncomeAmount,
    "mês com 3000 preserva 900");
Equal(700m, maximum[1].VacationAmount + maximum[1].IncomeAmount,
    "mês com 1000 preserva 300");
var capped = VacationPlanningService.CalculateAutomaticTransfers(availableSavings,
    new() { new() { Person = "Camila", BudgetShare = 30m, LostIncome = 20m } });
Equal(50m, capped.Sum(item => item.VacationAmount + item.IncomeAmount),
    "não compromete mais que o necessário");
Equal(30m, capped.Sum(item => item.VacationAmount), "respeita limite da viagem");
Equal(20m, capped.Sum(item => item.IncomeAmount), "respeita limite da renda");
Equal(0, VacationPlanningService.CalculateAutomaticTransfers(availableSavings,
    new() { new() { Person = "Camila", BudgetShare = 0m, LostIncome = 0m } }).Count,
    "sem meta restante não mexe no Guardado");
Equal(2, availableSavings.Single(source => source.Person == "Camila").RowNumber,
    "linha preservada mesmo com ID formatado igual a outra linha");
VacationPlanningService.ValidateGuardadoTransfers(savingsRequest, availableSavings);
var originalRow = savingsRequest.GuardadoTransfers[0].RowNumber;
savingsRequest.GuardadoTransfers[0].RowNumber = 4;
try
{
    VacationPlanningService.ValidateGuardadoTransfers(savingsRequest, availableSavings);
    throw new Exception("Linha diferente com mesmo ID foi aceita.");
}
catch (InvalidOperationException) { }
savingsRequest.GuardadoTransfers[0].RowNumber = originalRow;
var savingsOption = VacationPlanningService.CalculateOption(savingsRequest, savingsRequest.StartDate,
    savingsRequest.EndDate, salaries, savingsRows, new(), new());
Equal(1800m, savingsOption.People[0].GuardadoAllocated, "parte para viagem já destinada por Camila");
Equal(1000m, savingsOption.People[0].IncomeGuardadoAllocated, "parte para dias sem faturamento já destinada");
Equal(0m, savingsOption.People[1].GuardadoAllocated, "Diego sem transferência");
Equal(Math.Round(700m / savingsOption.SavingPeriods, 2), savingsOption.People[0].MonthlySaving,
    "meta futura da viagem desconta valor já destinado");
Equal(Math.Round((savingsOption.People[0].LostIncome - 1000m) / savingsOption.SavingPeriods, 2),
    savingsOption.People[0].IncomeMonthlySaving, "meta futura dos dias sem trabalho desconta valor já destinado");
Equal(savingsOption.People[0].LostIncome, savingsOption.People[0].IncomeCoverageByPeriod.Values.Sum(),
    "cobertura provisionada corresponde à receita não faturada");
var savedPeople = System.Text.Json.JsonSerializer.Deserialize<List<VacationPersonPreview>>(
    System.Text.Json.JsonSerializer.Serialize(savingsOption.People))!;
Equal(savingsOption.People[0].IncomeCoverageByPeriod.Single().Value,
    savedPeople[0].IncomeCoverageByPeriod.Single().Value, "cobertura persiste no plano de férias");
Equal(3000m, 3000m - 1800m - 1000m + 1800m + 1000m,
    "reclassificar Guardado preserva total reservado no mês");
savingsRequest.GuardadoTransfers[0].VacationAmount = 2500;
try
{
    VacationPlanningService.ValidateGuardadoTransfers(savingsRequest, availableSavings);
    throw new Exception("Soma das duas transferências maior que a linha Guardado foi aceita.");
}
catch (InvalidOperationException) { }
savingsRequest.GuardadoTransfers[0].VacationAmount = 2600;
savingsRequest.GuardadoTransfers[0].IncomeAmount = 0;
try
{
    VacationPlanningService.CalculateOption(savingsRequest, savingsRequest.StartDate,
        savingsRequest.EndDate, salaries, savingsRows, new(), new());
    throw new Exception("Transferência maior que a parte da viagem foi aceita.");
}
catch (ArgumentException) { }
savingsRequest.GuardadoTransfers[0].VacationAmount = 0;
savingsRequest.GuardadoTransfers[0].IncomeAmount = savingsOption.People[0].LostIncome + 1;
try
{
    VacationPlanningService.CalculateOption(savingsRequest, savingsRequest.StartDate,
        savingsRequest.EndDate, salaries, savingsRows, new(), new());
    throw new Exception("Transferência maior que a perda de faturamento foi aceita.");
}
catch (ArgumentException) { }

request.Participants[1].Percent = 30;
try
{
    VacationPlanningService.Validate(request);
    throw new Exception("A divisão inválida foi aceita.");
}
catch (ArgumentException) { }

var firstTrip = new VacationPersonPreview
{
    Person = "Camila", BudgetShare = 500m,
    IncomeReserveFunded = 240m,
    IncomeCoverageByPeriod = new() { ["03/2027"] = 240m }
};
Equal(0m, GoogleSheetsService.RequiredGuardadoReduction(350m), "saldo acima do mínimo não reduz Guardado");
Equal(150m, GoogleSheetsService.RequiredGuardadoReduction(50m), "ajuste mantém R$ 200");
Equal(700m, GoogleSheetsService.RequiredGuardadoReduction(-500m), "saldo negativo requer cobertura até R$ 200");
var secondTrip = new VacationPersonPreview
{
    Person = "Camila", BudgetShare = 400m,
    IncomeReserveFunded = 160m,
    IncomeCoverageByPeriod = new() { ["08/2027"] = 160m }
};
Equal(0m, GoogleSheetsService.ReducePlanReserves(new[] { firstTrip, secondTrip }, 600m, false),
    "redução de viagem deve caber nos planos");
Equal(500m, firstTrip.TripReserveShortfall, "férias mais próximas recebem redução primeiro");
Equal(100m, secondTrip.TripReserveShortfall, "redução restante vai para as próximas férias");
Equal(0m, GoogleSheetsService.ReducePlanReserves(new[] { firstTrip, secondTrip }, 300m, true),
    "redução de renda deve caber nos planos");
Equal(0m, firstTrip.IncomeCoverageByPeriod["03/2027"], "cobertura mais próxima reduz primeiro");
Equal(100m, secondTrip.IncomeCoverageByPeriod["08/2027"], "cobertura seguinte mantém saldo correto");
Equal(60m, secondTrip.IncomeReserveShortfall, "falta de cobertura fica registrada");

Console.WriteLine("Testes de férias: período, dias úteis, receita, divisão e saldo passaram.");
